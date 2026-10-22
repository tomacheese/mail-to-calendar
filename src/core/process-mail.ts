import type { Config } from '../config'
import {
  Graph,
  GraphError,
  type EventView,
  type Mail,
} from '../integrations/graph'
import { Foundry, FoundryError } from '../integrations/foundry'
import {
  State,
  failedAttempts,
  isSettled,
  sha,
  type BookingRow,
} from '../integrations/state'
import { classifyFailure, describeError } from './failure'
import {
  PendingPayloadSchema,
  validateBooking,
  type Analysis,
  type ValidBooking,
} from './validation'

/** One calendar operation decided (and possibly applied) for a booking. */
export interface Op {
  op: 'create' | 'update' | 'cancel' | 'skip' | 'hold'
  subject?: string
  start?: string
  end?: string
  note?: string
  applied: boolean
}
/** Result of one scan run. Contains booking details, so keep it out of production logs. */
export interface Summary {
  baseline: boolean
  seen: number
  processed: number
  ops: Op[]
  failures: number
  /** 1 when mail was left for the next run because the Foundry call budget was used up. */
  deferred: number
  foundryCalls: number
}

const norm = (s: string) =>
  s.normalize('NFKC').replaceAll(/\s+/g, '').toLowerCase()

/**
 * Fingerprint of the user-visible fields we manage. A mismatch between what we wrote and what Outlook holds
 * means the user edited the event (lastModifiedDateTime is useless: Outlook bumps it without content changes).
 */
export const fingerprint = (v: {
  subject: string
  start: string
  end: string
  location?: string
  showAs: string
  body: string
}) =>
  sha(
    [
      v.subject,
      v.start.slice(0, 19),
      v.end.slice(0, 19),
      v.location ?? '',
      v.showAs,
      norm(v.body).slice(0, 100),
    ].join('|')
  )
/**
 * Fingerprint of an event as Outlook currently holds it.
 */
const viewFingerprint = (cur: EventView) =>
  fingerprint({
    subject: cur.subject,
    start: cur.start.dateTime,
    end: cur.end.dateTime,
    location: cur.location?.displayName,
    showAs: cur.showAs,
    body: cur.bodyPreview,
  })

/** Builds the Graph event payload for a validated booking. */
export function renderEvent(b: ValidBooking, tentative: boolean) {
  const lines = [
    `施設: ${b.provider}`,
    b.location && `場所: ${b.location}`,
    b.reservation_number && `予約番号: ${b.reservation_number}`,
    b.details && `\n${b.details}`,
    '\n(予約メールから自動登録)',
  ].filter(Boolean)
  return {
    subject: `${tentative ? '【仮】' : ''}${b.subject}`,
    body: { contentType: 'text', content: lines.join('\n') },
    start: { dateTime: b.start, timeZone: b.timezone },
    end: { dateTime: b.end, timeZone: b.timezone },
    location: b.location ? { displayName: b.location } : undefined,
    showAs: tentative ? 'tentative' : 'busy',
    isReminderOn: !tentative,
    reminderMinutesBeforeStart: 30,
    categories: ['予約自動登録'],
  }
}

/**
 * Looks up the managed booking for a reservation.
 * Different providers can issue the same number, so a number alone never selects a row.
 */
async function findExisting(
  state: State,
  b: ValidBooking
): Promise<{
  row?: BookingRow
  key: string
  ambiguous?: boolean
  otherProvider?: boolean
}> {
  if (b.reservation_number) {
    const n = norm(b.reservation_number)
    const provider = norm(b.provider)
    const hits = await state.findByNumber(n)
    const same = hits.filter((h) => norm(h.provider) === provider)
    if (same.length > 1) return { key: '', ambiguous: true }
    return {
      row: same[0],
      key: same[0]?.key ?? sha(`${provider}|${n}`),
      otherProvider: same.length === 0 && hits.length > 0,
    }
  }
  const key = sha(`${norm(b.provider)}|${b.start}`)
  return { row: await state.getBooking(key), key }
}

type Found = Awaited<ReturnType<typeof findExisting>>

/**
 * Decides and applies the calendar operation for one booking, honouring CALENDAR_WRITE_MODE.
 * `receivedAt` is the receipt time of the mail the change comes from; it is stored so an older change is never applied over a newer one.
 */
async function handleBooking(
  cfg: Config,
  graph: Graph,
  state: State,
  action: Analysis['action'],
  b: ValidBooking,
  found: Found,
  receivedAt: Date
): Promise<Op> {
  if (found.ambiguous)
    return {
      op: 'hold',
      subject: b.subject,
      note: 'multiple bookings share this number',
      applied: false,
    }
  const mode = cfg.CALENDAR_WRITE_MODE
  const sourceReceivedAt = receivedAt.toISOString()
  const { row, key } = found
  const tentative =
    action === 'tentative' || (action === 'update' && !!row?.tentative)
  const base = { subject: b.subject, start: b.start, end: b.end }

  if (!row && (action === 'update' || action === 'cancel'))
    return {
      op: 'hold',
      ...base,
      note: found.otherProvider
        ? 'reservation number belongs to a different provider (not touching calendar)'
        : 'no managed booking matches (not touching calendar)',
      applied: false,
    }
  if (row?.status === 'detached')
    return {
      op: 'skip',
      ...base,
      note: 'user edited/deleted the event earlier; leaving it alone',
      applied: false,
    }

  const cancel = action === 'cancel'
  if (cancel && row?.status === 'cancelled')
    return {
      op: 'skip',
      ...base,
      note: 'already cancelled',
      applied: false,
    }
  if (!cancel && row?.status === 'cancelled')
    return {
      op: 'hold',
      ...base,
      note: 'booking was cancelled; not resurrecting',
      applied: false,
    }
  const ev = renderEvent(b, tentative)
  // Only schedule-relevant fields count as a change; the model rewords subject/details on every mail.
  const payloadHash = sha(`${b.start}|${b.end}|${b.timezone}|${tentative}`)
  if (row && !cancel && row.payloadHash === payloadHash)
    return {
      op: 'skip',
      ...base,
      note: 'duplicate of managed booking',
      applied: false,
    }
  if (row && !cancel && tentative && !row.tentative)
    return {
      op: 'skip',
      ...base,
      note: 'not downgrading a confirmed booking to tentative',
      applied: false,
    }

  const op: Op['op'] = row ? (cancel ? 'cancel' : 'update') : 'create'
  const allowed = mode === 'full' || (mode === 'create-only' && op === 'create')
  if (!allowed)
    return { op, ...base, note: `not applied (mode=${mode})`, applied: false }

  const number = b.reservation_number ? norm(b.reservation_number) : ''
  if (!row) {
    const ge = await graph.createEvent(ev, key.slice(0, 32))
    await state.putBooking({
      key,
      number,
      provider: b.provider,
      eventId: ge.id,
      subject: ev.subject,
      payloadHash,
      tentative,
      timezone: b.timezone,
      fingerprint: viewFingerprint(await graph.getEvent(ge.id, b.timezone)),
      status: 'active',
      sourceReceivedAt,
    })
    return { op, ...base, applied: true }
  }
  // Update/cancel: only if the event still looks exactly as we wrote it.
  let cur
  try {
    cur = await graph.getEvent(row.eventId, row.timezone)
  } catch (error) {
    if (error instanceof GraphError && error.status === 404) {
      await state.putBooking({ ...row, status: 'detached' })
      return {
        op: 'hold',
        ...base,
        note: 'event deleted by user',
        applied: false,
      }
    }
    throw error
  }
  if (viewFingerprint(cur) !== row.fingerprint) {
    await state.putBooking({ ...row, status: 'detached' })
    return { op: 'hold', ...base, note: 'event edited by user', applied: false }
  }
  if (cancel) {
    // Cancellation deletes the event (Deleted Items keeps it recoverable); the Bookings row stays as the audit record and blocks resurrection.
    await graph.deleteEvent(row.eventId)
    await state.putBooking({ ...row, status: 'cancelled', sourceReceivedAt })
    return { op, ...base, applied: true }
  }
  await graph.patchEvent(row.eventId, ev)
  await state.putBooking({
    ...row,
    subject: ev.subject,
    payloadHash,
    tentative,
    timezone: b.timezone,
    fingerprint: viewFingerprint(await graph.getEvent(row.eventId, b.timezone)),
    status: 'active',
    sourceReceivedAt,
  })
  return { op, ...base, applied: true }
}

/** Time an update/cancel that matched no booking stays queued before it is dropped. */
const PENDING_TTL_DAYS = 14

/** Failed apply attempts after which a pending item is given up. */
const PENDING_MAX_ATTEMPTS = 3

/** Per-field bounds that keep the stored payload far below the Azure Tables property limit. */
const PENDING_TEXT_MAX = 500
const PENDING_DETAILS_MAX = 4000
const PENDING_PAYLOAD_MAX = 30_000

const clamp = (s: string, max: number) => s.slice(0, max)
const clampOrNull = (s: string | null) =>
  s === null ? null : clamp(s, PENDING_TEXT_MAX)

/**
 * Queues an update/cancel that found no managed booking so it can be re-evaluated once the booking exists.
 * Text fields are bounded because a cancel skips the field limits of validateBooking.
 */
async function enqueueUnmatched(
  state: State,
  action: 'update' | 'cancel',
  b: ValidBooking,
  receivedAt: Date
): Promise<void> {
  const booking: ValidBooking = {
    ...b,
    provider: clamp(b.provider, PENDING_TEXT_MAX),
    subject: clamp(b.subject, PENDING_TEXT_MAX),
    location: clampOrNull(b.location),
    reservation_number: clampOrNull(b.reservation_number),
    timezone: clamp(b.timezone, PENDING_TEXT_MAX),
    target_reservation_hint: clampOrNull(b.target_reservation_hint),
    start_at: clampOrNull(b.start_at),
    end_at: clampOrNull(b.end_at),
    details: clamp(b.details, PENDING_DETAILS_MAX),
    evidence: [],
  }
  const excess =
    JSON.stringify({ action, booking }).length - PENDING_PAYLOAD_MAX
  if (excess > 0) {
    booking.details = booking.details.slice(
      0,
      Math.max(0, booking.details.length - excess)
    )
  }
  const number = booking.reservation_number
    ? norm(booking.reservation_number)
    : ''
  const matchKey = number || `${norm(booking.provider)}|${booking.start}`
  const now = new Date()
  await state.enqueuePending({
    key: sha(`${action}|${matchKey}|${receivedAt.toISOString()}`),
    status: 'open',
    action,
    number,
    receivedAt: receivedAt.toISOString(),
    payload: JSON.stringify({ action, booking }),
    reason: 'no managed booking matches',
    attempts: 0,
    createdAt: now.toISOString(),
    expiresAt: new Date(
      now.getTime() + PENDING_TTL_DAYS * 86_400e3
    ).toISOString(),
  })
}

/** Applies one analysis according to CALENDAR_WRITE_MODE and returns the operations performed or planned. */
export async function processMail(
  cfg: Config,
  graph: Graph,
  state: State,
  mail: Mail,
  analysis: Analysis
): Promise<Op[]> {
  if (analysis.action === 'ignore') return []
  if (analysis.action === 'needs_review')
    return [
      {
        op: 'hold',
        subject: mail.subject,
        note: `needs_review: ${analysis.reason}`,
        applied: false,
      },
    ]
  const ops: Op[] = []
  for (const original of analysis.bookings) {
    // A whitespace-only number would normalise to '' and match every number-less booking.
    const b = {
      ...original,
      reservation_number:
        original.reservation_number && norm(original.reservation_number)
          ? original.reservation_number
          : null,
    }
    // A cancellation often omits the date: a reservation number that really appears in the mail is enough to match it.
    const numberOnly =
      analysis.action === 'cancel' &&
      !b.start_at &&
      !!b.reservation_number &&
      norm(mail.bodyText).includes(norm(b.reservation_number))
    const v = numberOnly
      ? { ok: true as const, value: { ...b, start: '', end: '' } }
      : validateBooking(
          b,
          mail.receivedAt,
          mail.bodyText,
          cfg.DEFAULT_EVENT_DURATION_MINUTES,
          // Deleting an event does not use the model-written text, so a long field must not block a cancel.
          analysis.action !== 'cancel'
        )
    if (!v.ok) {
      ops.push({
        op: 'hold',
        subject: b.subject,
        note: `validation: ${v.reason}`,
        applied: false,
      })
      continue
    }
    const found = await findExisting(state, v.value)
    ops.push(
      await handleBooking(
        cfg,
        graph,
        state,
        analysis.action,
        v.value,
        found,
        mail.receivedAt
      )
    )
    // Only full mode can ever apply an update/cancel later; dry-run and create-only would just accumulate dead items.
    if (
      cfg.CALENDAR_WRITE_MODE === 'full' &&
      (analysis.action === 'update' || analysis.action === 'cancel') &&
      !found.row &&
      !found.ambiguous
    ) {
      await enqueueUnmatched(state, analysis.action, v.value, mail.receivedAt)
    }
  }
  return ops
}

/** Fetches a mail, or undefined when it no longer exists (404/410). */
async function fetchMail(graph: Graph, id: string): Promise<Mail | undefined> {
  try {
    return await graph.getMail(id)
  } catch (error) {
    if (error instanceof GraphError && [404, 410].includes(error.status)) {
      return undefined
    }
    throw error
  }
}

/** A mail that keeps failing for its own reasons is parked after this many attempts. */
const MAX_ATTEMPTS = 3

/** After this many mails parked in a row the cause is assumed to be systemic (for example a model change), so parking stops. */
const MAX_EXHAUSTED_STREAK = 2

/** A stored pending payload that cannot be read back; retrying cannot help. */
class PendingPayloadError extends Error {}

/** Parses a stored pending payload, rejecting anything that is not the shape enqueueUnmatched wrote. */
function parsePendingPayload(payload: string) {
  let raw: unknown
  try {
    raw = JSON.parse(payload)
  } catch {
    throw new PendingPayloadError('stored pending payload is not valid JSON')
  }
  const parsed = PendingPayloadSchema.safeParse(raw)
  if (!parsed.success) {
    throw new PendingPayloadError('stored pending payload has an invalid shape')
  }
  return parsed.data
}

/** Failures tied to one item (a bad payload or a Graph 4xx about this event) count as an attempt; anything else is an outage. */
function isPerItemFailure(error: unknown): boolean {
  return (
    error instanceof PendingPayloadError ||
    (error instanceof GraphError &&
      error.status >= 400 &&
      error.status < 500 &&
      ![401, 403, 408, 429].includes(error.status))
  )
}

/**
 * Re-evaluates queued update/cancel items against the bookings that exist now, oldest mail first.
 * An item is applied only when no newer mail has changed its booking; item-level problems are logged and never thrown.
 */
export async function reconcilePending(
  cfg: Config,
  graph: Graph,
  state: State,
  log: (m: string) => void,
  now = new Date()
): Promise<Op[]> {
  const items = await state.listOpenPending()
  const ops: Op[] = []
  const closed = { done: 0, superseded: 0, expired: 0, failed: 0 }
  for (const item of items) {
    const label = item.key.slice(0, 8)
    try {
      if (item.expiresAt <= now.toISOString()) {
        await state.updatePending(item.key, {
          status: 'expired',
          reason: 'expired before a matching booking appeared',
        })
        closed.expired++
        continue
      }
      const { action, booking } = parsePendingPayload(item.payload)
      const found = await findExisting(state, booking)
      if (found.ambiguous) {
        await state.updatePending(item.key, {
          status: 'done',
          reason: 'multiple bookings share this number',
        })
        closed.done++
        continue
      }
      if (!found.row) continue
      if (
        found.row.sourceReceivedAt !== undefined &&
        found.row.sourceReceivedAt > item.receivedAt
      ) {
        await state.updatePending(item.key, {
          status: 'superseded',
          reason: 'a newer mail already changed the booking',
        })
        closed.superseded++
        continue
      }
      const op = await handleBooking(
        cfg,
        graph,
        state,
        action,
        booking,
        found,
        new Date(item.receivedAt)
      )
      ops.push(op)
      await state.updatePending(item.key, {
        status: 'done',
        reason: op.note ?? 'applied',
      })
      closed.done++
    } catch (error) {
      log(`pending ${label} failed: ${describeError(error)}`)
      if (!isPerItemFailure(error)) break
      const attempts = item.attempts + 1
      const failed = attempts >= PENDING_MAX_ATTEMPTS
      try {
        await state.updatePending(
          item.key,
          failed
            ? { attempts, status: 'failed', reason: describeError(error) }
            : { attempts }
        )
      } catch (updateError) {
        log(`pending ${label} not updated: ${describeError(updateError)}`)
        break
      }
      if (failed) closed.failed++
    }
  }
  const open = items.length - Object.values(closed).reduce((a, n) => a + n, 0)
  log(
    `pending: open=${open} done=${closed.done} superseded=${closed.superseded} expired=${closed.expired} failed=${closed.failed}`
  )
  return ops
}

/**
 * Walks the Inbox delta, analyses new mail and applies the resulting operations.
 * Mail is handled oldest first and processing stops at the first failure that is not parked, so a cancel cannot overtake the booking it refers to.
 */
export async function runScan(
  cfg: Config,
  log: (m: string) => void = console.log
): Promise<Summary> {
  const graph = new Graph(cfg),
    foundry = new Foundry(cfg),
    state = new State(cfg)
  await state.init()
  log(
    `env=${cfg.APP_ENV} mode=${cfg.CALENDAR_WRITE_MODE} existing=${cfg.PROCESS_EXISTING_MESSAGES}`
  )

  const sync = await state.getSync()
  // Taken before the walk so mail arriving during it is never older than the baseline.
  const startedAt = new Date().toISOString()
  let page
  try {
    page = await graph.inboxDelta(sync?.deltaLink)
  } catch (error) {
    if (!sync || !(error instanceof GraphError) || error.status !== 410) {
      throw error
    }
    log('delta token expired (410); walking the Inbox again from scratch')
    page = await graph.inboxDelta()
  }
  const { messages, deltaLink } = page
  const initializedAt = sync?.initializedAt ?? startedAt
  const s: Summary = {
    baseline: !sync,
    seen: messages.length,
    processed: 0,
    ops: [],
    failures: 0,
    deferred: 0,
    foundryCalls: 0,
  }

  if (!sync && !cfg.PROCESS_EXISTING_MESSAGES) {
    await state.setSync(deltaLink, initializedAt) // baseline only: existing mail is never turned into events
    return s
  }

  const dryRun = cfg.CALENDAR_WRITE_MODE === 'dry-run'
  let streak = await state.getExhaustedStreak()
  messages.sort((a, b) =>
    (a.receivedDateTime ?? '').localeCompare(b.receivedDateTime ?? '')
  )
  for (const m of messages) {
    if (m.removed || m.isDraft) continue
    // Delta also reports read/move changes on old mail; only handle mail received after baseline.
    if (
      !cfg.PROCESS_EXISTING_MESSAGES &&
      m.receivedDateTime &&
      m.receivedDateTime < initializedAt
    )
      continue
    const prior = await state.getProcessedResult(m.id)
    if (isSettled(prior, dryRun)) continue
    // Hitting the call budget is throttling, not a failure: the rest waits for the next run.
    if (foundry.exhausted) {
      s.deferred++
      break
    }
    const receivedAt = new Date(m.receivedDateTime ?? Date.now())
    const label = sha(m.id).slice(0, 8)
    const prefix = dryRun ? 'dry-run:' : ''
    try {
      const mail = await fetchMail(graph, m.id)
      if (!mail) {
        // Moved or deleted since the delta reported it: nothing left to process.
        await state.markProcessed(m.id, `${prefix}error:gone`, receivedAt)
        log(`message ${label} is gone`)
        continue
      }
      const analysis = await foundry.analyze(mail)
      const ops = await processMail(cfg, graph, state, mail, analysis)
      s.ops.push(...ops)
      // Dry-run keeps the cursor back (so write modes can replay) but remembers the plan to avoid re-inference every tick.
      if (dryRun)
        await state.markProcessed(
          m.id,
          `dry-run:${analysis.action}`,
          mail.receivedAt,
          JSON.stringify(
            ops.map(({ op, subject, start, end, note }) => ({
              op,
              subject,
              start,
              end,
              note,
            }))
          )
        )
      else await state.markProcessed(m.id, analysis.action, mail.receivedAt)
      s.processed++
      if (streak > 0) {
        streak = 0
        await state.setExhaustedStreak(0)
      }
    } catch (error) {
      if (error instanceof FoundryError && error.permanent) {
        // Content filter and similar: retrying cannot help, so park it and let the cursor advance.
        await state.markProcessed(
          m.id,
          `${prefix}error:${error.status}`,
          receivedAt
        )
        log(`message ${label} parked: ${error.message}`)
        continue
      }
      s.failures++
      log(`message ${label} failed: ${describeError(error)}`)
      if (classifyFailure(error) === 'per-mail') {
        const attempts = failedAttempts(prior) + 1
        if (attempts >= MAX_ATTEMPTS && streak < MAX_EXHAUSTED_STREAK) {
          await state.markProcessed(
            m.id,
            `${prefix}error:exhausted`,
            receivedAt
          )
          streak++
          await state.setExhaustedStreak(streak)
          log(`message ${label} parked after ${attempts} attempts`)
          continue
        }
        if (attempts >= MAX_ATTEMPTS) {
          log(
            `${streak} mails in a row were parked; treating this as systemic and holding the rest`
          )
        }
        await state.markProcessed(m.id, `retry:${attempts}`, receivedAt)
      }
      break
    }
  }
  s.foundryCalls = foundry.calls
  // Needs neither Foundry nor the delta, so it also runs when the loop stopped early; its failures never hold the cursor.
  if (!dryRun && cfg.CALENDAR_WRITE_MODE === 'full') {
    try {
      s.ops.push(...(await reconcilePending(cfg, graph, state, log)))
    } catch (error) {
      log(`pending reconcile failed: ${describeError(error)}`)
    }
  }
  // Advance the cursor only when every message is settled; dry-run never advances it.
  if (!dryRun && s.failures === 0 && s.deferred === 0)
    await state.setSync(deltaLink, initializedAt)
  return s
}
