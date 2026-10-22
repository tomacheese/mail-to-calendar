import { createHash } from 'node:crypto'
import { TableClient } from '@azure/data-tables'
import type { Config } from '../config'

/** Managed booking: reservation identity, the Outlook event it maps to, and the fingerprint of what we last wrote. */
export interface BookingRow {
  key: string
  number: string
  provider: string
  eventId: string
  subject: string
  payloadHash: string
  tentative: boolean
  timezone: string
  fingerprint: string
  status: 'active' | 'cancelled' | 'detached'
  /** Receipt time of the latest mail whose change was applied; rows written before this field existed have none. */
  sourceReceivedAt?: string
}
/** An update/cancel that matched no managed booking when its mail was handled, kept until the booking appears or the item closes. */
export interface PendingRow {
  key: string
  status: 'open' | 'done' | 'superseded' | 'expired' | 'failed'
  action: 'update' | 'cancel'
  number: string
  receivedAt: string
  payload: string
  reason: string
  attempts: number
  createdAt: string
  expiresAt: string
}
export const sha = (s: string) => createHash('sha256').update(s).digest('hex')

/** Persistent state on Azure Tables: sync cursor, processed mails, managed bookings and pending update/cancel items. */
export class State {
  private t: Record<
    'SyncState' | 'ProcessedMessages' | 'Bookings' | 'PendingOps',
    TableClient
  >
  constructor(cfg: Config) {
    const mk = (n: string) =>
      TableClient.fromConnectionString(
        cfg.STATE_STORAGE_CONNECTION,
        cfg.STATE_TABLE_PREFIX + n,
        { allowInsecureConnection: true }
      )
    this.t = {
      SyncState: mk('SyncState'),
      ProcessedMessages: mk('ProcessedMessages'),
      Bookings: mk('Bookings'),
      PendingOps: mk('PendingOps'),
    }
  }
  /** Creates the tables if they do not exist. */
  async init() {
    await Promise.all(Object.values(this.t).map((c) => c.createTable()))
  }

  private async get<T extends object>(
    t: TableClient,
    pk: string,
    rk: string
  ): Promise<T | undefined> {
    try {
      return await t.getEntity<T>(pk, rk)
    } catch (error) {
      if ((error as { statusCode?: number }).statusCode === 404)
        return undefined
      throw error
    }
  }

  /** Returns the saved delta cursor and the time the baseline was taken. */
  getSync() {
    return this.get<{ deltaLink: string; initializedAt: string }>(
      this.t.SyncState,
      'sync',
      'inbox'
    )
  }
  /** Number of mails parked as exhausted in a row; reset by any mail that succeeds. */
  async getExhaustedStreak(): Promise<number> {
    const row = await this.get<{ count: number }>(
      this.t.SyncState,
      'sync',
      'exhausted'
    )
    return row?.count ?? 0
  }
  /** Saves the exhausted-in-a-row counter. */
  setExhaustedStreak(count: number) {
    return this.t.SyncState.upsertEntity({
      partitionKey: 'sync',
      rowKey: 'exhausted',
      count,
    })
  }
  /** Saves the delta cursor. */
  setSync(deltaLink: string, initializedAt: string) {
    return this.t.SyncState.upsertEntity({
      partitionKey: 'sync',
      rowKey: 'inbox',
      deltaLink,
      initializedAt,
    })
  }

  /** Returns the recorded result of a mail (`<action>`, `dry-run:*`, `error:*` or `retry:<n>`), if any. */
  async getProcessedResult(id: string): Promise<string | undefined> {
    const r = await this.get<{ result: string }>(
      this.t.ProcessedMessages,
      'm',
      sha(id)
    )
    return r?.result
  }
  /** Records that a mail was handled; plan holds the dry-run operations. */
  markProcessed(id: string, result: string, receivedAt: Date, plan = '') {
    return this.t.ProcessedMessages.upsertEntity({
      partitionKey: 'm',
      rowKey: sha(id),
      result,
      receivedAt: receivedAt.toISOString(),
      plan,
    })
  }
  /** Lists plans remembered by dry-run runs, oldest first. */
  async listDryRunPlans(): Promise<
    { receivedAt: string; result: string; plan: string }[]
  > {
    const out = []
    const entities = this.t.ProcessedMessages.listEntities<{
      receivedAt: string
      result: string
      plan: string
    }>({ queryOptions: { filter: "PartitionKey eq 'm'" } })
    for await (const e of entities) {
      if (e.result.startsWith('dry-run'))
        out.push({ receivedAt: e.receivedAt, result: e.result, plan: e.plan })
    }
    return out.toSorted((a, b) => a.receivedAt.localeCompare(b.receivedAt))
  }

  /** Looks up a managed booking by key. */
  async getBooking(key: string): Promise<BookingRow | undefined> {
    const e = await this.get<BookingRow>(this.t.Bookings, 'b', key)
    return e && { ...e, key }
  }
  /** Bookings sharing a reservation number (providers may differ in spelling, so match on number only). */
  async findByNumber(number: string): Promise<BookingRow[]> {
    const out: BookingRow[] = []
    const entities = this.t.Bookings.listEntities<
      BookingRow & { rowKey: string }
    >({
      queryOptions: { filter: `number eq '${number.replaceAll("'", "''")}'` },
    })
    for await (const e of entities) {
      out.push({ ...e, key: e.rowKey })
    }
    return out
  }
  /** Inserts or replaces a managed booking. */
  putBooking(b: BookingRow) {
    const { key, ...rest } = b
    return this.t.Bookings.upsertEntity({
      partitionKey: 'b',
      rowKey: key,
      ...rest,
    })
  }

  /** Stores a pending item unless one with the same key exists; returns false for that duplicate so replays are harmless. */
  async enqueuePending(row: PendingRow): Promise<boolean> {
    const { key, ...rest } = row
    try {
      await this.t.PendingOps.createEntity({
        partitionKey: 'p',
        rowKey: key,
        ...rest,
      })
      return true
    } catch (error) {
      if ((error as { statusCode?: number }).statusCode === 409) return false
      throw error
    }
  }
  /** Open pending items, oldest mail first so a change is never applied before an earlier one. */
  async listOpenPending(): Promise<PendingRow[]> {
    const out: PendingRow[] = []
    const entities = this.t.PendingOps.listEntities<
      Omit<PendingRow, 'key'> & { rowKey: string }
    >({
      queryOptions: { filter: "PartitionKey eq 'p' and status eq 'open'" },
    })
    for await (const e of entities) {
      out.push({
        key: e.rowKey,
        status: e.status,
        action: e.action,
        number: e.number,
        receivedAt: e.receivedAt,
        payload: e.payload,
        reason: e.reason,
        attempts: e.attempts,
        createdAt: e.createdAt,
        expiresAt: e.expiresAt,
      })
    }
    return out.toSorted((a, b) => a.receivedAt.localeCompare(b.receivedAt))
  }
  /** Updates the outcome fields of a pending item (last writer wins). */
  updatePending(
    key: string,
    patch: Partial<Pick<PendingRow, 'status' | 'reason' | 'attempts'>>
  ) {
    return this.t.PendingOps.updateEntity(
      { partitionKey: 'p', rowKey: key, ...patch },
      'Merge'
    )
  }
}

/**
 * Whether a recorded result settles the mail. A `dry-run:*` mark only suppresses re-analysis in dry-run, and `retry:<n>` never settles it.
 */
export function isSettled(
  result: string | undefined,
  dryRun: boolean
): boolean {
  const pending = result === undefined || result.startsWith('retry:')
  return !pending && (dryRun || !result.startsWith('dry-run'))
}

/** Number of failed attempts recorded for a mail. */
export function failedAttempts(result: string | undefined): number {
  return result?.startsWith('retry:')
    ? Number(result.slice('retry:'.length))
    : 0
}
