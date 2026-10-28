import { randomUUID } from 'node:crypto'
import { TableClient } from '@azure/data-tables'
import type { Config } from '../src/config'
import type { PendingRow } from '../src/integrations/state'

const table = (cfg: Config, name: string) =>
  TableClient.fromConnectionString(
    cfg.STATE_STORAGE_CONNECTION,
    cfg.STATE_TABLE_PREFIX + name,
    { allowInsecureConnection: true }
  )

/** Every pending item in the config's table, oldest mail first, including closed ones (State only lists open ones). */
export async function pendingRows(cfg: Config) {
  const out: {
    status: string
    action: string
    number: string
    receivedAt: string
    reason: string
    attempts: number
  }[] = []
  const entities = table(cfg, 'PendingOps').listEntities<
    Omit<PendingRow, 'key'>
  >()
  for await (const e of entities) {
    out.push({
      status: e.status,
      action: e.action,
      number: e.number,
      receivedAt: e.receivedAt,
      reason: e.reason,
      attempts: e.attempts,
    })
  }
  return out.toSorted((a, b) => a.receivedAt.localeCompare(b.receivedAt))
}

/** Rewrites a booking row as an older version of the app would have: without sourceReceivedAt. */
export async function stripSourceReceivedAt(
  cfg: Config,
  key: string
): Promise<void> {
  const t = table(cfg, 'Bookings')
  const e = await t.getEntity<
    { partitionKey: string; rowKey: string } & Record<string, unknown>
  >('b', key)
  const {
    sourceReceivedAt: _removed,
    etag: _etag,
    timestamp: _timestamp,
    'odata.metadata': _metadata,
    ...rest
  } = e
  await t.upsertEntity(rest, 'Replace')
}

/** Payload as enqueueUnmatched stores it; empty start/end is a number-only cancel. */
export function pendingPayload(
  action: 'update' | 'cancel',
  number: string,
  when: { start: string; end: string } = { start: '', end: '' },
  provider = 'テスト歯科'
): string {
  return JSON.stringify({
    action,
    booking: {
      provider,
      subject: '検診',
      start_at: null,
      end_at: null,
      timezone: 'Asia/Tokyo',
      location: null,
      reservation_number: number,
      details: 'd',
      evidence: [],
      target_reservation_hint: null,
      ...when,
    },
  })
}

/** A pending item that has not expired, for seeding State directly. */
export function pendingItem(
  number: string,
  receivedAt: Date,
  overrides: Partial<PendingRow> = {}
): PendingRow {
  return {
    key: randomUUID(),
    status: 'open',
    action: 'cancel',
    number,
    receivedAt: receivedAt.toISOString(),
    payload: pendingPayload('cancel', number),
    reason: 'no managed booking matches',
    attempts: 0,
    createdAt: new Date().toISOString(),
    expiresAt: new Date(Date.now() + 14 * 86_400e3).toISOString(),
    ...overrides,
  }
}
