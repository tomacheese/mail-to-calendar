import type { Config } from '../config'
import {
  Graph,
  GraphError,
  type EventView,
  type Mail,
} from '../integrations/graph'
import { Foundry, FoundryError } from '../integrations/foundry'
import { State, sha, type BookingRow } from '../integrations/state'
import { validateBooking, type Analysis, type ValidBooking } from './validation'

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

/** Existing managed booking for this reservation, or undefined. Number match wins; otherwise provider+start. */
async function findExisting(
  state: State,
  b: ValidBooking
): Promise<{ row?: BookingRow; key: string; ambiguous?: boolean }> {
  if (b.reservation_number) {
    const n = norm(b.reservation_number)
    const hits = await state.findByNumber(n)
    if (hits.length > 1) return { key: '', ambiguous: true }
    return {
      row: hits[0],
      key: hits[0]?.key ?? sha(`${norm(b.provider)}|${n}`),
    }
  }
  const key = sha(`${norm(b.provider)}|${b.start}`)
  return { row: await state.getBooking(key), key }
}

/** Decides and applies the calendar operation for one booking, honouring CALENDAR_WRITE_MODE. */
async function handleBooking(
  cfg: Config,
  graph: Graph,
  state: State,
  action: Analysis['action'],
  b: ValidBooking
): Promise<Op> {
  const mode = cfg.CALENDAR_WRITE_MODE
  const tentative = action === 'tentative'
  const found = await findExisting(state, b)
  if (found.ambiguous)
    return {
      op: 'hold',
      subject: b.subject,
      note: 'multiple bookings share this number',
      applied: false,
    }
  const { row, key } = found
  const base = { subject: b.subject, start: b.start, end: b.end }

  if (!row && (action === 'update' || action === 'cancel'))
    return {
      op: 'hold',
      ...base,
      note: 'no managed booking matches (not touching calendar)',
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
    await state.putBooking({ ...row, status: 'cancelled' })
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
  })
  return { op, ...base, applied: true }
}

/** Pure decision on what one analysis means; returns ops. */
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
  for (const b of analysis.bookings) {
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
          cfg.DEFAULT_EVENT_DURATION_MINUTES
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
    ops.push(await handleBooking(cfg, graph, state, analysis.action, v.value))
  }
  return ops
}

/** Walks the Inbox delta, analyses new mail and applies the resulting operations. */
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
  const { messages, deltaLink } = await graph.inboxDelta(sync?.deltaLink)
  const initializedAt = sync?.initializedAt ?? new Date().toISOString()
  const s: Summary = {
    baseline: !sync,
    seen: messages.length,
    processed: 0,
    ops: [],
    failures: 0,
    foundryCalls: 0,
  }

  if (!sync && !cfg.PROCESS_EXISTING_MESSAGES) {
    await state.setSync(deltaLink, initializedAt) // baseline only: existing mail is never turned into events
    return s
  }

  const dryRun = cfg.CALENDAR_WRITE_MODE === 'dry-run'
  // Delta order is not chronological; a cancel must never be handled before the booking it refers to.
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
    if (await state.isProcessed(m.id, dryRun)) continue
    try {
      const mail = await graph.getMail(m.id)
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
    } catch (error) {
      if (error instanceof FoundryError && error.permanent) {
        // Retrying cannot help (e.g. content filter); park it so the cursor can advance.
        await state.markProcessed(
          m.id,
          `${dryRun ? 'dry-run:' : ''}error:${error.status}`,
          new Date(m.receivedDateTime ?? Date.now())
        )
        log(`message parked: ${error.message}`)
        continue
      }
      s.failures++
      log(`message failed: ${(error as Error).message}`)
    }
  }
  s.foundryCalls = foundry.calls
  // Advance the cursor only when every message is settled; dry-run never advances it.
  if (s.failures === 0 && cfg.CALENDAR_WRITE_MODE !== 'dry-run')
    await state.setSync(deltaLink, initializedAt)
  return s
}
