import { z } from 'zod'

export const ACTIONS = [
  'confirmed',
  'tentative',
  'update',
  'cancel',
  'ignore',
  'needs_review',
] as const
export type Action = (typeof ACTIONS)[number]

export const BookingSchema = z.object({
  provider: z.string(),
  subject: z.string(),
  // Local wall-clock time without offset, e.g. 2026-10-20T14:30:00; zone is in `timezone`.
  start_at: z.string().nullable(),
  end_at: z.string().nullable(),
  timezone: z.string(),
  location: z.string().nullable(),
  reservation_number: z.string().nullable(),
  details: z.string(),
  evidence: z.array(z.string()),
  target_reservation_hint: z.string().nullable(),
})

export const AnalysisSchema = z.object({
  action: z.enum(ACTIONS),
  bookings: z.array(BookingSchema),
  reason: z.string(),
})

export type Booking = z.infer<typeof BookingSchema>
export type Analysis = z.infer<typeof AnalysisSchema>

/** Stored form of an update/cancel that found no managed booking yet (a number-only cancel keeps empty start/end). */
export const PendingPayloadSchema = z.object({
  action: z.enum(['update', 'cancel']),
  booking: BookingSchema.extend({ start: z.string(), end: z.string() }),
})

/** A booking that passed validation, with normalised local start/end. */
export interface ValidBooking extends Booking {
  start: string // local datetime, no offset, seconds precision
  end: string
}

const LOCAL_DT = /^(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d)(?::(\d\d))?$/
const ABS_DATE =
  /\d{1,2}月\d{1,2}日|\d{4}[/.-]\d{1,2}[/.-]\d{1,2}|\d{1,2}\/\d{1,2}/
const norm = (s: string) => s.normalize('NFKC').replaceAll(/\s+/g, '')

/** Upper bounds for model-written text that ends up in the calendar; Graph rejects over-long subjects. */
const FIELD_LIMITS = [
  ['subject', 200],
  ['provider', 200],
  ['location', 500],
  ['details', 4000],
] as const

/** Parse a local datetime as if it were UTC (wall-clock arithmetic only). */
function wall(s: string): number | null {
  const m = LOCAL_DT.exec(s)
  if (!m) return null
  const [, year, month, day, hour, minute, second = '0'] = m
  const t = Date.UTC(
    Number(year),
    Number(month) - 1,
    Number(day),
    Number(hour),
    Number(minute),
    Number(second)
  )
  const back = new Date(t)
  // Reject rollovers such as 2026-02-31.
  return back.getUTCMonth() === Number(month) - 1 &&
    back.getUTCDate() === Number(day)
    ? t
    : null
}

const iso = (t: number) => new Date(t).toISOString().slice(0, 19)

/**
 * Deterministic gate between model output and any calendar write.
 * Never trusts the model: format, timezone, plausibility and evidence-in-body are re-checked here.
 */
export function validateBooking(
  b: Booking,
  receivedAt: Date,
  bodyText: string,
  defaultMinutes: number,
  enforceLimits = true
): { ok: true; value: ValidBooking } | { ok: false; reason: string } {
  for (const [field, max] of FIELD_LIMITS) {
    if (enforceLimits && (b[field]?.length ?? 0) > max) {
      return { ok: false, reason: `${field} is longer than ${max} characters` }
    }
  }
  if (!b.start_at) return { ok: false, reason: 'start_at missing' }
  const s = wall(b.start_at)
  if (s === null)
    return { ok: false, reason: 'start_at not a valid local datetime' }
  try {
    new Intl.DateTimeFormat('en', { timeZone: b.timezone })
  } catch {
    return { ok: false, reason: 'invalid timezone' }
  }
  const r = receivedAt.getTime()
  // Wall-clock vs UTC differs by at most ~14h; 36h slack keeps this offset-agnostic.
  if (s < r - 36 * 3600e3)
    return { ok: false, reason: 'start_at is in the past relative to receipt' }
  if (s > r + 2 * 365 * 86_400e3)
    return { ok: false, reason: 'start_at implausibly far in the future' }

  const body = norm(bodyText)
  const found = b.evidence
    .map((e) => norm(e))
    .filter((e) => e && body.includes(e))
  if (found.every((e) => !/\d/.test(e)))
    return { ok: false, reason: 'no date/time evidence found in mail body' }
  // Relative wording ("来週火曜") is ambiguous for the model too; require an explicit calendar date in the quoted evidence.
  if (found.every((e) => !ABS_DATE.test(e)))
    return { ok: false, reason: 'no explicit calendar date in evidence' }

  let e = b.end_at ? wall(b.end_at) : null
  if (e === null || e <= s) e = s + defaultMinutes * 60e3
  return { ok: true, value: { ...b, start: iso(s), end: iso(e) } }
}
