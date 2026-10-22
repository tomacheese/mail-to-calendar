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
  /** Per-event log (`create:ev1`, `patch:ev1`, `delete:ev2`) for telling which event an operation hit. */
  const events: string[] = []
  const store = new Map<string, any>()
  let last = ''
  const view = (ev: any, cur: any) => ({
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
      last = `ev${store.size + 1}`
      store.set(last, view(ev, {}))
      events.push(`create:${last}`)
      return { id: last, lastModifiedDateTime: 'x', subject: '' }
    },
    getEvent: async (id: string) => ({ id, ...store.get(id) }),
    // Bumping lastModifiedDateTime alone must NOT look like a user edit, so it is not even modelled here.
    deleteEvent: async (id: string) => {
      calls.push('delete')
      events.push(`delete:${id}`)
    },
    patchEvent: async (id: string, p: any) => {
      calls.push('patch')
      events.push(`patch:${id}`)
      store.set(id, view(p, store.get(id)))
      return { id, lastModifiedDateTime: 'y', subject: '' }
    },
  } as unknown as Graph
  return {
    g,
    calls,
    events,
    userEdit: () => {
      const cur = store.get(last)
      store.set(last, { ...cur, subject: cur.subject + ' (my note)' })
    },
  }
}

const num = `N-${randomUUID().slice(0, 8)}`
const mk = (
  action: Analysis['action'],
  start: string,
  n = num,
  provider = 'テスト歯科'
): Analysis => ({
  action,
  reason: '',
  bookings: [
    {
      provider,
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

describe('施設ごとの予約の同一性', () => {
  const providerB = '別の医院'
  const differentProviderNote =
    'reservation number belongs to a different provider (not touching calendar)'
  // Rows store the normalised (lower-cased) number, which findByNumber expects.
  const newNumber = () => `n-${randomUUID().slice(0, 8)}`
  const full = cfg('full')

  it('別の施設から同じ予約番号の新規予約が届いても、既存の予定を上書きせず別の予約として作成する', async () => {
    const state = new State(full)
    await state.init()
    const { g, calls } = stub()
    const n = newNumber(),
      t = `${soon(5)}T10:00:00`
    await processMail(full, g, state, mail(t), mk('confirmed', t, n))
    const [before] = await state.findByNumber(n)

    const result = await processMail(
      full,
      g,
      state,
      mail(t),
      mk('confirmed', t, n, providerB)
    )

    expect(result[0]).toMatchObject({ op: 'create', applied: true })
    expect(calls).toEqual(['create', 'create'])
    const rows = await state.findByNumber(n)
    expect(rows).toHaveLength(2)
    expect(new Set(rows.map((r) => r.key)).size).toBe(2)
    expect(rows.map((r) => r.provider).toSorted()).toEqual(
      ['テスト歯科', providerB].toSorted()
    )
    const after = rows.find((r) => r.provider === 'テスト歯科')
    expect(after).toMatchObject({
      key: before!.key,
      eventId: before!.eventId,
      payloadHash: before!.payloadHash,
      status: 'active',
    })
  })

  it('別の施設から同じ予約番号の変更が届いても、既存の予定は更新せず保留する', async () => {
    const state = new State(full)
    await state.init()
    const { g, calls } = stub()
    const n = newNumber(),
      t1 = `${soon(5)}T10:00:00`,
      t2 = `${soon(6)}T11:00:00`
    await processMail(full, g, state, mail(t1), mk('confirmed', t1, n))
    const [before] = await state.findByNumber(n)

    const result = await processMail(
      full,
      g,
      state,
      mail(t2),
      mk('update', t2, n, providerB)
    )

    expect(result[0]).toMatchObject({
      op: 'hold',
      note: differentProviderNote,
      applied: false,
    })
    expect(calls).toEqual(['create'])
    expect(await state.findByNumber(n)).toEqual([before])
  })

  it('別の施設から同じ予約番号の取消が届いても、既存の予定は削除せず保留する', async () => {
    const state = new State(full)
    await state.init()
    const { g, calls } = stub()
    const n = newNumber(),
      t = `${soon(5)}T10:00:00`
    await processMail(full, g, state, mail(t), mk('confirmed', t, n))

    const result = await processMail(
      full,
      g,
      state,
      mail(t),
      mk('cancel', t, n, providerB)
    )

    expect(result[0]).toMatchObject({
      op: 'hold',
      note: differentProviderNote,
      applied: false,
    })
    expect(calls).toEqual(['create'])
    const [row] = await state.findByNumber(n)
    expect(row?.status).toBe('active')
  })

  it('日付のない取消でも、別の施設の予約番号では既存の予定を削除せず保留する', async () => {
    const state = new State(full)
    await state.init()
    const { g, calls } = stub()
    const n = newNumber(),
      t = `${soon(5)}T10:00:00`
    await processMail(full, g, state, mail(t), mk('confirmed', t, n))
    const cancel = mk('cancel', t, n, providerB)
    cancel.bookings[0]!.start_at = null
    cancel.bookings[0]!.evidence = []

    const result = await processMail(
      full,
      g,
      state,
      { ...mail(t), bodyText: `予約番号 ${n} をキャンセルしました` },
      cancel
    )

    expect(result[0]).toMatchObject({
      op: 'hold',
      note: differentProviderNote,
      applied: false,
    })
    expect(calls).toEqual(['create'])
    const [row] = await state.findByNumber(n)
    expect(row?.status).toBe('active')
  })

  it('同じ予約番号の予約が複数の施設にあっても、変更と取消は一致した施設の予定だけに作用する', async () => {
    const state = new State(full)
    await state.init()
    const { g, events } = stub()
    const n = newNumber(),
      t1 = `${soon(5)}T10:00:00`,
      t2 = `${soon(6)}T11:00:00`,
      t3 = `${soon(7)}T12:00:00`
    await processMail(full, g, state, mail(t1), mk('confirmed', t1, n))
    await processMail(
      full,
      g,
      state,
      mail(t2),
      mk('confirmed', t2, n, providerB)
    )
    const rowB = (await state.findByNumber(n)).find(
      (r) => r.provider === providerB
    )
    expect(events).toEqual(['create:ev1', 'create:ev2'])

    const updated = await processMail(
      full,
      g,
      state,
      mail(t3),
      mk('update', t3, n)
    )
    const cancelled = await processMail(
      full,
      g,
      state,
      mail(t3),
      mk('cancel', t3, n)
    )

    expect(updated[0]).toMatchObject({ op: 'update', applied: true })
    expect(cancelled[0]).toMatchObject({ op: 'cancel', applied: true })
    expect(events).toEqual([
      'create:ev1',
      'create:ev2',
      'patch:ev1',
      'delete:ev1',
    ])
    const afterB = (await state.findByNumber(n)).find(
      (r) => r.provider === providerB
    )
    expect(afterB).toEqual(rowB)
  })

  it('同じ施設の表記ゆれ（全角・空白・大文字小文字）は同じ予約として更新する', async () => {
    const state = new State(full)
    await state.init()
    const { g, calls } = stub()
    const n = newNumber(),
      t1 = `${soon(5)}T10:00:00`,
      t2 = `${soon(6)}T11:00:00`
    await processMail(
      full,
      g,
      state,
      mail(t1),
      mk('confirmed', t1, n, 'Test Clinic')
    )

    const result = await processMail(
      full,
      g,
      state,
      mail(t2),
      mk('update', t2, n, 'ＴＥＳＴ  clinic')
    )

    expect(result[0]).toMatchObject({ op: 'update', applied: true })
    expect(calls).toEqual(['create', 'patch'])
    expect(await state.findByNumber(n)).toHaveLength(1)
  })

  it('同じ施設・同じ予約番号の行が複数あるときは、何も書き込まず保留する', async () => {
    const state = new State(full)
    await state.init()
    const { g, calls } = stub()
    const n = newNumber(),
      t = `${soon(5)}T10:00:00`
    for (const key of [`${n}-1`, `${n}-2`]) {
      await state.putBooking({
        key,
        number: n,
        provider: 'テスト歯科',
        eventId: `event-${key}`,
        subject: '検診',
        payloadHash: 'h',
        tentative: false,
        timezone: 'Asia/Tokyo',
        fingerprint: 'f',
        status: 'active',
      })
    }

    const result = await processMail(
      full,
      g,
      state,
      mail(t),
      mk('update', t, n)
    )

    expect(result[0]).toMatchObject({
      op: 'hold',
      note: 'multiple bookings share this number',
      applied: false,
    })
    expect(calls).toEqual([])
  })

  it('どの施設にも該当する予約がない変更と取消は、従来どおりの理由で保留する', async () => {
    const state = new State(full)
    await state.init()
    const { g, calls } = stub()
    const t = `${soon(5)}T10:00:00`
    const noBooking = 'no managed booking matches (not touching calendar)'

    for (const action of ['update', 'cancel'] as const) {
      const result = await processMail(
        full,
        g,
        state,
        mail(t),
        mk(action, t, newNumber())
      )
      expect(result[0]).toMatchObject({
        op: 'hold',
        note: noBooking,
        applied: false,
      })
    }
    expect(calls).toEqual([])
  })
})
