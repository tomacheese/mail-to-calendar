// Network is stubbed via global fetch (Graph, Entra, Foundry); state uses Azurite with an isolated table prefix.
import { randomUUID } from 'node:crypto'
import { loadConfig } from '../src/config'
import { runScan } from '../src/core/process-mail'
import { Graph } from '../src/integrations/graph'
import { Foundry } from '../src/integrations/foundry'

const env = () => ({
  GRAPH_TENANT_ID: randomUUID(),
  GRAPH_CLIENT_ID: randomUUID(),
  GRAPH_CLIENT_SECRET: 's3cr3t-value-'.padEnd(30, 'x'),
  GRAPH_MAILBOX: 't@example.com',
  FOUNDRY_ENDPOINT: 'https://foundry.example.com',
  FOUNDRY_DEPLOYMENT: 'd',
  FOUNDRY_API_KEY: 'key-12345',
  STATE_STORAGE_CONNECTION: process.env.TEST_STATE_CONNECTION!,
  STATE_TABLE_PREFIX: 'T' + randomUUID().replace(/-/g, '').slice(0, 10),
})
const json = (b: unknown, status = 200, h: Record<string, string> = {}) =>
  new Response(JSON.stringify(b), { status, headers: h })

const realFetch = globalThis.fetch
afterEach(() => {
  globalThis.fetch = realFetch
})

// ---- config ----
it('config: invalid mode and swapped credentials are rejected without echoing values', () => {
  expect(() => loadConfig({ ...env(), CALENDAR_WRITE_MODE: 'yolo' })).toThrow()
  const e = env()
  expect(() =>
    loadConfig({
      ...e,
      GRAPH_CLIENT_ID: 'abc~secretvalue_that_is_not_a_guid',
      GRAPH_CLIENT_SECRET: randomUUID(),
    })
  ).toThrow(/^(?![\s\S]*secretvalue)/) // 値はエラーに含めない
  expect(loadConfig(env()).CALENDAR_WRITE_MODE).toBe('dry-run') // 未設定は dry-run に倒す
})

// ---- Graph retry / error hygiene ----
it('graph: retries 429 then succeeds; errors carry status+code only', async () => {
  let n = 0
  globalThis.fetch = (async (u: string) => {
    if (String(u).includes('login.'))
      return json({ access_token: 't', expires_in: 3600 })
    if (++n === 1) return json({}, 429, { 'retry-after': '0' })
    return json({ totalItemCount: 7 })
  }) as typeof fetch
  expect(await new Graph(loadConfig(env())).inboxCount()).toBe(7)
  expect(n).toBe(2)

  globalThis.fetch = (async (u: string) =>
    String(u).includes('login.')
      ? json({ access_token: 't' })
      : json(
          {
            error: {
              code: 'ErrorAccessDenied',
              message: 'SECRET BODY mail content',
            },
          },
          403
        )) as typeof fetch
  await expect(new Graph(loadConfig(env())).inboxCount()).rejects.toThrow(
    'Graph GET 403 ErrorAccessDenied'
  )
})

// ---- Foundry cap / error hygiene ----
it('foundry: call cap enforced; 4xx is permanent and leaks no body', async () => {
  const cfg = loadConfig({ ...env(), FOUNDRY_MAX_CALLS_PER_RUN: '1' })
  const ok = {
    choices: [
      {
        message: {
          content: JSON.stringify({
            action: 'ignore',
            bookings: [],
            reason: 'x',
          }),
        },
      },
    ],
  }
  globalThis.fetch = (async () => json(ok)) as typeof fetch
  const f = new Foundry(cfg),
    m = { subject: 's', from: 'f', receivedAt: new Date(), bodyText: 'b' }
  await f.analyze(m)
  await expect(f.analyze(m)).rejects.toThrow(/MAX_CALLS/)

  globalThis.fetch = (async () =>
    json({ error: { message: 'PROMPT ECHO' } }, 400)) as typeof fetch
  await expect(new Foundry(loadConfig(env())).analyze(m)).rejects.toMatchObject(
    {
      permanent: true,
      message: 'Foundry 400',
    }
  )
})

// ---- runScan: cursor, replay, idempotency ----
function scenario() {
  const s = {
    delta: [] as { id: string; receivedDateTime: string }[],
    nextLink: 0,
    foundryFail: false,
    foundryStatus: 200,
    creates: 0,
    lastEvent: {} as Record<string, any>,
    link: 1,
  }
  globalThis.fetch = (async (u: string, init?: RequestInit) => {
    const url = String(u)
    if (url.includes('login.'))
      return json({ access_token: 't', expires_in: 3600 })
    if (url.includes('openai/v1')) {
      if (s.foundryFail) throw new Error('network down')
      if (s.foundryStatus !== 200) return json({}, s.foundryStatus)
      const start = new Date(Date.now() + 5 * 86400e3)
      const d = start.toISOString().slice(0, 10),
        md = `${start.getUTCMonth() + 1}月${start.getUTCDate()}日 10:00`
      return json({
        choices: [
          {
            message: {
              content: JSON.stringify({
                action: 'confirmed',
                reason: '',
                bookings: [
                  {
                    provider: 'P',
                    subject: 'S',
                    start_at: `${d}T10:00:00`,
                    end_at: null,
                    timezone: 'Asia/Tokyo',
                    location: null,
                    reservation_number: 'Z' + s.link,
                    details: '',
                    evidence: [md],
                    target_reservation_hint: null,
                  },
                ],
              }),
            },
          },
        ],
      })
    }
    if (url.includes('/messages/delta') || url.includes('deltatoken')) {
      const value = s.delta.splice(0)
      return json({
        value,
        '@odata.deltaLink': `https://graph.microsoft.com/v1.0/x/messages/delta?deltatoken=${++s.link}`,
      })
    }
    if (url.includes('/events/') && init?.method === 'GET') {
      const e = s.lastEvent
      return json({
        id: 'ev' + s.creates,
        subject: e.subject,
        showAs: e.showAs,
        bodyPreview: e.body.content,
        start: e.start,
        end: e.end,
        location: e.location,
      })
    }
    if (url.includes('/events') && init?.method === 'POST') {
      s.creates++
      s.lastEvent = JSON.parse(init.body as string)
      return json(
        {
          id: 'ev' + s.creates,
          lastModifiedDateTime: '2026-10-18T00:00:00Z',
          subject: '',
        },
        201
      )
    }
    if (url.includes('/messages/')) {
      const d = new Date(Date.now() + 5 * 86400e3)
      return json({
        subject: '予約',
        receivedDateTime: new Date().toISOString(),
        from: { emailAddress: { address: 'a@b.c' } },
        body: { content: `${d.getUTCMonth() + 1}月${d.getUTCDate()}日 10:00` },
      })
    }
    throw new Error('unexpected ' + url)
  }) as typeof fetch
  return s
}
const quiet = () => {}

it('runScan: baseline → new mail → failure keeps cursor → retry succeeds → no reprocessing', async () => {
  const s = scenario()
  const cfg = { ...loadConfig(env()), CALENDAR_WRITE_MODE: 'full' as const }

  const base = await runScan(cfg, quiet)
  expect(base).toMatchObject({ baseline: true, processed: 0 })
  expect(s.creates).toBe(0) // ベースラインでは予定を作らない

  s.delta = [
    { id: 'm1', receivedDateTime: new Date(Date.now() + 1000).toISOString() },
  ]
  s.foundryFail = true
  const failed = await runScan(cfg, quiet)
  expect(failed.failures).toBe(1)

  s.foundryFail = false
  s.delta = [
    { id: 'm1', receivedDateTime: new Date(Date.now() + 1000).toISOString() },
  ] // delta replays m1 because the cursor did not advance
  const ok = await runScan(cfg, quiet)
  expect(ok).toMatchObject({ failures: 0, processed: 1 })
  expect(s.creates).toBe(1)

  s.delta = [{ id: 'm1', receivedDateTime: new Date().toISOString() }] // read-state change re-reports m1
  const again = await runScan(cfg, quiet)
  expect(again.processed).toBe(0) // 処理済みメールは二重処理しない
  expect(s.creates).toBe(1)
})

it('runScan: permanent Foundry 4xx parks the message and lets the cursor advance', async () => {
  const s = scenario()
  const cfg = { ...loadConfig(env()), CALENDAR_WRITE_MODE: 'full' as const }
  await runScan(cfg, quiet)
  s.delta = [
    { id: 'bad', receivedDateTime: new Date(Date.now() + 1000).toISOString() },
  ]
  s.foundryStatus = 400
  const r = await runScan(cfg, quiet)
  expect(r.failures).toBe(0)
  s.foundryStatus = 200
  s.delta = [{ id: 'bad', receivedDateTime: new Date().toISOString() }]
  expect((await runScan(cfg, quiet)).processed).toBe(0) // 保留したメールは再試行しない
})

it('runScan: removed/draft/old messages are ignored; dry-run never marks processed or advances cursor', async () => {
  const s = scenario()
  const cfg = loadConfig(env()) // dry-run
  await runScan(cfg, quiet)
  const future = new Date(Date.now() + 1000).toISOString()
  s.delta = [
    { id: 'old', receivedDateTime: '2020-01-01T00:00:00Z' },
    { id: 'd', receivedDateTime: future, isDraft: true } as never,
    { id: 'r', '@removed': {} } as never,
  ]
  const r = await runScan(cfg, quiet)
  expect(r.processed).toBe(0)
  s.delta = [{ id: 'n', receivedDateTime: future }]
  expect((await runScan(cfg, quiet)).processed).toBe(1)
  s.delta = [{ id: 'n', receivedDateTime: future }]
  expect((await runScan(cfg, quiet)).processed).toBe(0) // dry-run は計画を記憶し、再推論しない
  // Switching to a write mode replays it: the dry-run mark must not block real processing.
  s.delta = [{ id: 'n', receivedDateTime: future }]
  const full = await runScan({ ...cfg, CALENDAR_WRITE_MODE: 'full' }, quiet)
  expect(full.processed).toBe(1)
  expect(s.creates).toBe(1)
})
