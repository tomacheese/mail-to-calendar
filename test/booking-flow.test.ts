// Azurite is started by test/global-setup.ts on a free port. Graph is stubbed; nothing leaves the machine.
import { randomUUID } from 'node:crypto'
import { processMail } from '../src/core/process-mail'
import { State } from '../src/integrations/state'
import { loadConfig } from '../src/config'
import type { Analysis } from '../src/core/validation'
import type { Graph } from '../src/integrations/graph'

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
const cfg = (mode: string) => loadConfig({ ...env, CALENDAR_WRITE_MODE: mode })

function stub() {
  const calls: string[] = []
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
    // Bumping lastModifiedDateTime alone must NOT look like a user edit, so it is not even modelled here.
    deleteEvent: async () => {
      calls.push('delete')
    },
    patchEvent: async (_id: string, p: any) => {
      calls.push('patch')
      cur = view(p)
      return { id: 'ev1', lastModifiedDateTime: 'y', subject: '' }
    },
  } as unknown as Graph
  return {
    g,
    calls,
    userEdit: () => {
      cur = { ...cur, subject: cur.subject + ' (my note)' }
    },
  }
}

const num = `N-${randomUUID().slice(0, 8)}`
const mk = (action: Analysis['action'], start: string, n = num): Analysis => ({
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
      evidence: [
        `${Number(start.slice(5, 7))}月${Number(start.slice(8, 10))}日 ${start.slice(11, 16)}`,
      ],
      target_reservation_hint: null,
    },
  ],
})
const soon = (d: number) =>
  new Date(Date.now() + d * 86400e3).toISOString().slice(0, 10)
const mail = (start: string) => ({
  subject: 's',
  from: 'f',
  receivedAt: new Date(),
  bodyText: `${Number(start.slice(5, 7))}月${Number(start.slice(8, 10))}日 ${start.slice(11, 16)}`,
})

it('create → duplicate skip → update → user edit detaches → cancel is held', async () => {
  const state = new State(cfg('full'))
  await state.init()
  const { g, calls, userEdit } = stub()
  const t1 = `${soon(5)}T10:00:00`,
    t2 = `${soon(6)}T11:00:00`
  const run = (a: Analysis, t: string, mode = 'full') =>
    processMail(cfg(mode), g, state, mail(t), a)

  expect((await run(mk('confirmed', t1), t1))[0]?.op).toBe('create')
  expect((await run(mk('confirmed', t1), t1))[0]?.op).toBe('skip')
  expect((await run(mk('update', t2), t2))[0]?.op).toBe('update')
  expect(calls).toEqual(['create', 'patch'])

  userEdit()
  const held = await run(mk('update', t1), t1)
  expect(held[0]?.op).toBe('hold')
  expect(calls).toHaveLength(2) // 手動編集された予定は更新しない
  expect((await run(mk('cancel', t1), t1))[0]?.op).toBe('skip') // 管理対象外になった予約は触らない
})

it('unknown cancel/update is held; dry-run and create-only never write existing', async () => {
  const state = new State(cfg('full'))
  await state.init()
  const { g, calls } = stub()
  const n = `N-${randomUUID().slice(0, 8)}`,
    t = `${soon(5)}T10:00:00`
  expect(
    (await processMail(cfg('full'), g, state, mail(t), mk('cancel', t, n)))[0]
      ?.op
  ).toBe('hold')

  const dry = await processMail(
    cfg('dry-run'),
    g,
    state,
    mail(t),
    mk('confirmed', t, n)
  )
  expect(dry[0]).toMatchObject({ op: 'create', applied: false })
  expect(calls).toHaveLength(0)

  await processMail(
    cfg('create-only'),
    g,
    state,
    mail(t),
    mk('confirmed', t, n)
  ) // creates
  const t2 = `${soon(7)}T09:00:00`
  const upd = await processMail(
    cfg('create-only'),
    g,
    state,
    mail(t2),
    mk('update', t2, n)
  )
  expect(upd[0]?.applied).toBe(false)
  expect(calls).toEqual(['create'])
})

it('cancel without a date matches by reservation number; without a number it is held', async () => {
  const state = new State(cfg('full'))
  await state.init()
  const { g, calls } = stub()
  const n = `N-${randomUUID().slice(0, 8)}`,
    t = `${soon(5)}T10:00:00`
  await processMail(cfg('full'), g, state, mail(t), mk('confirmed', t, n))
  const noDate = (num: string | null): Analysis => ({
    action: 'cancel',
    reason: '',
    bookings: [
      {
        ...mk('cancel', t, n).bookings[0]!,
        start_at: null,
        evidence: [],
        reservation_number: num,
      },
    ],
  })
  const body = { ...mail(t), bodyText: `予約番号 ${n} をキャンセルしました` }
  expect(
    (await processMail(cfg('full'), g, state, body, noDate(null)))[0]?.op
  ).toBe('hold')
  expect(
    (
      await processMail(
        cfg('full'),
        g,
        state,
        { ...body, bodyText: '別の予約をキャンセル' },
        noDate(n)
      )
    )[0]?.op
  ).toBe('hold') // 予約番号は本文に実在する必要がある
  expect(
    (await processMail(cfg('full'), g, state, body, noDate(n)))[0]?.op
  ).toBe('cancel')
  expect(calls).toEqual(['create', 'delete'])
  expect(
    (
      await processMail(cfg('full'), g, state, mail(t), mk('confirmed', t, n))
    )[0]?.op
  ).toBe('hold') // 取消済みは復活させない
})

it('更新の後でも、手を加えていない予定は取消で削除できる', async () => {
  const state = new State(cfg('full'))
  await state.init()
  const { g, calls } = stub()
  const n = `N-${randomUUID().slice(0, 8)}`
  const t1 = `${soon(5)}T10:00:00`
  const t2 = `${soon(6)}T11:00:00`
  await processMail(cfg('full'), g, state, mail(t1), mk('confirmed', t1, n))
  await processMail(cfg('full'), g, state, mail(t2), mk('update', t2, n))
  const result = await processMail(
    cfg('full'),
    g,
    state,
    mail(t2),
    mk('cancel', t2, n)
  )
  expect(result[0]).toMatchObject({ op: 'cancel', applied: true })
  expect(calls).toEqual(['create', 'patch', 'delete'])
})

it('場所の無い更新メールの後でも、手を加えていない予定は取消で削除できる', async () => {
  const state = new State(cfg('full'))
  await state.init()
  const { g } = stub()
  const n = `N-${randomUUID().slice(0, 8)}`
  const t1 = `${soon(5)}T10:00:00`
  const t2 = `${soon(6)}T11:00:00`
  const withLocation = mk('confirmed', t1, n)
  withLocation.bookings[0]!.location = '東京都港区1-1-1'
  await processMail(cfg('full'), g, state, mail(t1), withLocation)
  await processMail(cfg('full'), g, state, mail(t2), mk('update', t2, n))
  const result = await processMail(
    cfg('full'),
    g,
    state,
    mail(t2),
    mk('cancel', t2, n)
  )
  expect(result[0]).toMatchObject({ op: 'cancel', applied: true })
})

describe('状態遷移の補強', () => {
  it('取消済みの予約に再度取消が来ても、状態を変えず削除も繰り返さない', async () => {
    const state = new State(cfg('full'))
    await state.init()
    const { g, calls } = stub()
    const n = `N-${randomUUID().slice(0, 8)}`
    const t = `${soon(5)}T10:00:00`
    await processMail(cfg('full'), g, state, mail(t), mk('confirmed', t, n))
    await processMail(cfg('full'), g, state, mail(t), mk('cancel', t, n))
    const again = await processMail(
      cfg('full'),
      g,
      state,
      mail(t),
      mk('cancel', t, n)
    )
    expect(again[0]).toMatchObject({ op: 'skip', note: 'already cancelled' })
    expect(calls).toEqual(['create', 'delete'])
    const [row] = await state.findByNumber(n.toLowerCase())
    expect(row?.status).toBe('cancelled')
  })

  it('update は仮予約を確定に昇格させない', async () => {
    const state = new State(cfg('full'))
    await state.init()
    const { g } = stub()
    const n = `N-${randomUUID().slice(0, 8)}`
    const t1 = `${soon(5)}T10:00:00`
    const t2 = `${soon(6)}T11:00:00`
    await processMail(cfg('full'), g, state, mail(t1), mk('tentative', t1, n))
    await processMail(cfg('full'), g, state, mail(t2), mk('update', t2, n))
    const [row] = await state.findByNumber(n.toLowerCase())
    expect(row?.tentative).toBe(true)
  })

  it('空白だけの予約番号は番号なしとして扱い、既存の予約に一致させない', async () => {
    const state = new State(cfg('full'))
    await state.init()
    const { g, calls } = stub()
    const t1 = `${soon(5)}T10:00:00`
    const t2 = `${soon(6)}T11:00:00`
    const first = mk('confirmed', t1, '')
    first.bookings[0]!.provider = `P-${randomUUID().slice(0, 8)}`
    await processMail(cfg('full'), g, state, mail(t1), first)
    const blank = mk('cancel', t2, '   ')
    blank.bookings[0]!.start_at = null
    blank.bookings[0]!.evidence = []
    const result = await processMail(
      cfg('full'),
      g,
      state,
      { ...mail(t2), bodyText: 'キャンセルしました' },
      blank
    )
    expect(result[0]?.op).toBe('hold')
    expect(calls).toEqual(['create'])
  })
})
