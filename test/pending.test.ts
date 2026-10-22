// Azurite is started by test/global-setup.ts on a free port. Graph is stubbed; nothing leaves the machine.
import { randomUUID } from 'node:crypto'
import { processMail, reconcilePending } from '../src/core/process-mail'
import { State } from '../src/integrations/state'
import { loadConfig } from '../src/config'
import { GraphError, type Graph } from '../src/integrations/graph'
import type { Analysis } from '../src/core/validation'
import {
  pendingItem,
  pendingRows,
  stripSourceReceivedAt,
} from './pending-helpers'

const env = {
  GRAPH_TENANT_ID: randomUUID(),
  GRAPH_CLIENT_ID: randomUUID(),
  GRAPH_CLIENT_SECRET: 'x'.repeat(30),
  GRAPH_MAILBOX: 't@example.com',
  FOUNDRY_ENDPOINT: 'https://example.com',
  FOUNDRY_DEPLOYMENT: 'd',
  FOUNDRY_API_KEY: 'k',
  STATE_STORAGE_CONNECTION: process.env.TEST_STATE_CONNECTION!,
}
const quiet = () => {}

/** Config, State and Graph stub on tables no other test touches, because reconcile handles every open item in its tables. */
async function setup() {
  const prefix = 'T' + randomUUID().replaceAll('-', '').slice(0, 10)
  const cfg = (mode = 'full') =>
    loadConfig({
      ...env,
      STATE_TABLE_PREFIX: prefix,
      CALENDAR_WRITE_MODE: mode,
    })
  const state = new State(cfg())
  await state.init()
  const calls: string[] = []
  const fail: { delete?: Error; patch?: Error } = {}
  const tries = { delete: 0 }
  let cur: any = {}
  const view = (ev: any) => ({
    subject: ev.subject ?? cur.subject,
    showAs: ev.showAs ?? cur.showAs,
    bodyPreview: ev.body?.content ?? cur.bodyPreview,
    start: ev.start ?? cur.start,
    end: ev.end ?? cur.end,
    location: ev.location ?? cur.location,
  })
  const g = {
    createEvent: async (ev: any) => {
      calls.push('create')
      cur = view(ev)
      return { id: 'ev1', lastModifiedDateTime: 'x', subject: '' }
    },
    getEvent: async () => ({ id: 'ev1', ...cur }),
    deleteEvent: async () => {
      tries.delete++
      if (fail.delete) throw fail.delete
      calls.push('delete')
    },
    patchEvent: async (_id: string, p: any) => {
      if (fail.patch) throw fail.patch
      calls.push('patch')
      cur = view(p)
      return { id: 'ev1', lastModifiedDateTime: 'y', subject: '' }
    },
  } as unknown as Graph
  return {
    cfg,
    state,
    g,
    calls,
    fail,
    tries,
    userEdit: () => {
      cur = { ...cur, subject: cur.subject + ' (my note)' }
    },
  }
}

const T0 = Date.now()
const at = (ms: number) => new Date(T0 + ms)
const soon = (d: number) =>
  new Date(Date.now() + d * 86400e3).toISOString().slice(0, 10)
const number = () => `N-${randomUUID().slice(0, 8)}`
const day = (start: string) =>
  `${Number(start.slice(5, 7))}月${Number(start.slice(8, 10))}日 ${start.slice(11, 16)}`

const mk = (
  action: Analysis['action'],
  start: string,
  n: string
): Analysis => ({
  action,
  reason: '',
  bookings: [
    {
      provider: 'テスト歯科',
      subject: '検診',
      start_at: start,
      end_at: null,
      timezone: 'Asia/Tokyo',
      location: null,
      reservation_number: n,
      details: 'd',
      evidence: [day(start)],
      target_reservation_hint: null,
    },
  ],
})
const mail = (receivedAt: Date, start: string) => ({
  subject: 's',
  from: 'f',
  receivedAt,
  bodyText: day(start),
})

describe('update/cancel が先に届いた場合', () => {
  it('取消は保留に積まれ、予約が作られた次の再評価で反映される', async () => {
    const { cfg, state, g, calls } = await setup()
    const n = number()
    const t = `${soon(5)}T10:00:00`

    const held = await processMail(
      cfg(),
      g,
      state,
      mail(at(2000), t),
      mk('cancel', t, n)
    )
    expect(held[0]).toMatchObject({ op: 'hold', applied: false })
    expect(calls).toEqual([])
    expect(await pendingRows(cfg())).toEqual([
      {
        status: 'open',
        action: 'cancel',
        number: n.toLowerCase(),
        receivedAt: at(2000).toISOString(),
        reason: 'no managed booking matches',
        attempts: 0,
      },
    ])

    await processMail(cfg(), g, state, mail(at(1000), t), mk('confirmed', t, n))
    const ops = await reconcilePending(cfg(), g, state, quiet)
    expect(ops).toHaveLength(1)
    expect(ops[0]).toMatchObject({ op: 'cancel', applied: true })
    expect(calls).toEqual(['create', 'delete'])
    expect((await state.findByNumber(n.toLowerCase()))[0]?.status).toBe(
      'cancelled'
    )
    expect(await pendingRows(cfg())).toMatchObject([
      { status: 'done', reason: 'applied', attempts: 0 },
    ])
  })

  it('変更は保留に積まれ、予約が作られた次の再評価で反映される', async () => {
    const { cfg, state, g, calls } = await setup()
    const n = number()
    const t1 = `${soon(5)}T10:00:00`
    const t2 = `${soon(6)}T11:00:00`

    await processMail(cfg(), g, state, mail(at(2000), t2), mk('update', t2, n))
    await processMail(
      cfg(),
      g,
      state,
      mail(at(1000), t1),
      mk('confirmed', t1, n)
    )
    const ops = await reconcilePending(cfg(), g, state, quiet)
    expect(ops[0]).toMatchObject({ op: 'update', applied: true })
    expect(calls).toEqual(['create', 'patch'])
    expect(await pendingRows(cfg())).toMatchObject([
      { status: 'done', action: 'update' },
    ])
  })

  it('一致する予約がない間は保留のまま変わらず、期限を過ぎると何も書かずに失効する', async () => {
    const { cfg, state, g, calls } = await setup()
    const n = number()
    const t = `${soon(5)}T10:00:00`
    await processMail(cfg(), g, state, mail(at(1000), t), mk('cancel', t, n))

    expect(await reconcilePending(cfg(), g, state, quiet)).toEqual([])
    expect(await reconcilePending(cfg(), g, state, quiet)).toEqual([])
    expect(await pendingRows(cfg())).toMatchObject([
      { status: 'open', attempts: 0 },
    ])

    const later = new Date(Date.now() + 15 * 86_400e3)
    expect(await reconcilePending(cfg(), g, state, quiet, later)).toEqual([])
    expect(await pendingRows(cfg())).toMatchObject([
      { status: 'expired', attempts: 0 },
    ])
    expect(calls).toEqual([])
  })

  it('解決済みの項目は再実行しても何も書かず、同じメールを二度積んでも1件のまま', async () => {
    const { cfg, state, g, calls } = await setup()
    const n = number()
    const t = `${soon(5)}T10:00:00`
    const cancelMail = mail(at(2000), t)
    await processMail(cfg(), g, state, cancelMail, mk('cancel', t, n))
    await processMail(cfg(), g, state, cancelMail, mk('cancel', t, n))
    expect(await pendingRows(cfg())).toHaveLength(1)

    await processMail(cfg(), g, state, mail(at(1000), t), mk('confirmed', t, n))
    expect(await reconcilePending(cfg(), g, state, quiet)).toHaveLength(1)
    expect(await reconcilePending(cfg(), g, state, quiet)).toEqual([])
    expect(calls).toEqual(['create', 'delete'])
  })

  it('予約がより新しいメールで作られていれば、古い取消は書き込まずに破棄する', async () => {
    const { cfg, state, g, calls } = await setup()
    const n = number()
    const t = `${soon(5)}T10:00:00`
    await processMail(cfg(), g, state, mail(at(1000), t), mk('cancel', t, n))
    await processMail(cfg(), g, state, mail(at(2000), t), mk('confirmed', t, n))

    expect(await reconcilePending(cfg(), g, state, quiet)).toEqual([])
    expect(calls).toEqual(['create'])
    expect((await state.findByNumber(n.toLowerCase()))[0]?.status).toBe(
      'active'
    )
    expect(await pendingRows(cfg())).toMatchObject([
      {
        status: 'superseded',
        reason: 'a newer mail already changed the booking',
      },
    ])
  })

  it('受信時刻を持たない旧形式の予約は新しいメール扱いにせず、取消を反映する', async () => {
    const { cfg, state, g, calls } = await setup()
    const n = number()
    const t = `${soon(5)}T10:00:00`
    await processMail(cfg(), g, state, mail(at(1000), t), mk('cancel', t, n))
    await processMail(cfg(), g, state, mail(at(2000), t), mk('confirmed', t, n))
    const [row] = await state.findByNumber(n.toLowerCase())
    await stripSourceReceivedAt(cfg(), row!.key)
    expect((await state.findByNumber(n.toLowerCase()))[0]).not.toHaveProperty(
      'sourceReceivedAt'
    )

    const ops = await reconcilePending(cfg(), g, state, quiet)
    expect(ops[0]).toMatchObject({ op: 'cancel', applied: true })
    expect(calls).toEqual(['create', 'delete'])
  })
})

describe('反映に失敗した場合', () => {
  /** A booking plus a queued cancel for it, ready for reconcile. */
  async function queued(t: Awaited<ReturnType<typeof setup>>, n: string) {
    const start = `${soon(5)}T10:00:00`
    await processMail(
      t.cfg(),
      t.g,
      t.state,
      mail(at(2000), start),
      mk('cancel', start, n)
    )
    await processMail(
      t.cfg(),
      t.g,
      t.state,
      mail(at(1000), start),
      mk('confirmed', start, n)
    )
  }

  it('項目固有の Graph エラー(400)は試行として数え、3回目で failed にしてもスキャンは落とさない', async () => {
    const t = await setup()
    await queued(t, number())
    t.fail.delete = new GraphError(400, 'Graph DELETE 400 ErrorInvalidRequest')

    for (const attempts of [1, 2]) {
      expect(await reconcilePending(t.cfg(), t.g, t.state, quiet)).toEqual([])
      expect(await pendingRows(t.cfg())).toMatchObject([
        { status: 'open', attempts },
      ])
    }
    expect(await reconcilePending(t.cfg(), t.g, t.state, quiet)).toEqual([])
    expect(await pendingRows(t.cfg())).toMatchObject([
      {
        status: 'failed',
        attempts: 3,
        reason: 'Graph DELETE 400 ErrorInvalidRequest',
      },
    ])
    expect(await reconcilePending(t.cfg(), t.g, t.state, quiet)).toEqual([])
    expect(t.calls).toEqual(['create'])
  })

  it('障害に当たるエラー(503)は試行に数えず、その回の残りの項目に進まない', async () => {
    const t = await setup()
    const n = number()
    const start = `${soon(5)}T10:00:00`
    for (const ms of [2000, 3000]) {
      await processMail(
        t.cfg(),
        t.g,
        t.state,
        mail(at(ms), start),
        mk('cancel', start, n)
      )
    }
    await processMail(
      t.cfg(),
      t.g,
      t.state,
      mail(at(1000), start),
      mk('confirmed', start, n)
    )
    t.fail.delete = new GraphError(503, 'Graph DELETE 503 ErrorServerBusy')

    expect(await reconcilePending(t.cfg(), t.g, t.state, quiet)).toEqual([])
    expect(t.tries.delete).toBe(1)
    expect(await pendingRows(t.cfg())).toMatchObject([
      { status: 'open', attempts: 0 },
      { status: 'open', attempts: 0 },
    ])

    delete t.fail.delete
    const ops = await reconcilePending(t.cfg(), t.g, t.state, quiet)
    expect(ops).toMatchObject([
      { op: 'cancel', applied: true },
      { op: 'skip', note: 'already cancelled' },
    ])
    expect(t.calls).toEqual(['create', 'delete'])
  })

  it('読み戻せない保存内容は Graph を呼ばずに試行として数え、failed になる', async () => {
    const t = await setup()
    await t.state.enqueuePending(
      pendingItem(number(), at(1000), { payload: '{}' })
    )
    for (const attempts of [1, 2]) {
      await reconcilePending(t.cfg(), t.g, t.state, quiet)
      expect(await pendingRows(t.cfg())).toMatchObject([
        { status: 'open', attempts },
      ])
    }
    await reconcilePending(t.cfg(), t.g, t.state, quiet)
    expect(await pendingRows(t.cfg())).toMatchObject([
      {
        status: 'failed',
        attempts: 3,
        reason: 'stored pending payload has an invalid shape',
      },
    ])
    expect(t.calls).toEqual([])
  })
})

describe('書き込みを行わない結果になる場合', () => {
  const booking = (n: string, key: string) => ({
    key,
    number: n.toLowerCase(),
    provider: 'テスト歯科',
    eventId: 'ev1',
    subject: '検診',
    payloadHash: 'h',
    tentative: false,
    timezone: 'Asia/Tokyo',
    fingerprint: 'f',
    status: 'active' as const,
  })

  it('同じ番号の予約が複数ある場合は何もせず完了にする', async () => {
    const { cfg, state, g, calls } = await setup()
    const n = number()
    await state.enqueuePending(pendingItem(n, at(1000)))
    await state.putBooking(booking(n, 'k1'))
    await state.putBooking(booking(n, 'k2'))

    expect(await reconcilePending(cfg(), g, state, quiet)).toEqual([])
    expect(await pendingRows(cfg())).toMatchObject([
      { status: 'done', reason: 'multiple bookings share this number' },
    ])
    expect(calls).toEqual([])
  })

  it('手動編集で管理対象外になった予約は触らず完了にする', async () => {
    const { cfg, state, g, calls } = await setup()
    const n = number()
    await state.putBooking({ ...booking(n, 'k1'), status: 'detached' })
    await state.enqueuePending(pendingItem(n, at(1000)))

    const ops = await reconcilePending(cfg(), g, state, quiet)
    expect(ops[0]).toMatchObject({ op: 'skip', applied: false })
    expect(await pendingRows(cfg())).toMatchObject([
      {
        status: 'done',
        reason: 'user edited/deleted the event earlier; leaving it alone',
      },
    ])
    expect(calls).toEqual([])
  })

  it('取消済みの予約への取消は何もせず完了にする', async () => {
    const { cfg, state, g, calls } = await setup()
    const n = number()
    await state.putBooking({ ...booking(n, 'k1'), status: 'cancelled' })
    await state.enqueuePending(pendingItem(n, at(1000)))

    await reconcilePending(cfg(), g, state, quiet)
    expect(await pendingRows(cfg())).toMatchObject([
      { status: 'done', reason: 'already cancelled' },
    ])
    expect(calls).toEqual([])
  })

  it('予定が手動編集されていれば削除せず、完了にする', async () => {
    const { cfg, state, g, calls, userEdit } = await setup()
    const n = number()
    const t = `${soon(5)}T10:00:00`
    await processMail(cfg(), g, state, mail(at(2000), t), mk('cancel', t, n))
    await processMail(cfg(), g, state, mail(at(1000), t), mk('confirmed', t, n))
    userEdit()

    await reconcilePending(cfg(), g, state, quiet)
    expect(await pendingRows(cfg())).toMatchObject([
      { status: 'done', reason: 'event edited by user' },
    ])
    expect(calls).toEqual(['create'])
  })
})

describe('保留に積まないモード', () => {
  it.each(['dry-run', 'create-only'])(
    '%s では一致しない取消を積まない',
    async (mode) => {
      const { cfg, state, g, calls } = await setup()
      const t = `${soon(5)}T10:00:00`
      const ops = await processMail(
        cfg(mode),
        g,
        state,
        mail(at(1000), t),
        mk('cancel', t, number())
      )
      expect(ops[0]).toMatchObject({ op: 'hold', applied: false })
      expect(await pendingRows(cfg())).toEqual([])
      expect(calls).toEqual([])
    }
  )
})
