import { validateBooking, type Booking } from '../src/core/validation'

const base: Booking = {
  provider: 'テスト歯科',
  subject: '歯科検診',
  start_at: '2026-10-20T14:30:00',
  end_at: null,
  timezone: 'Asia/Tokyo',
  location: null,
  reservation_number: 'A-1',
  details: '',
  evidence: ['10月20日 14:30'],
  target_reservation_hint: null,
}
const body = 'ご予約日時： 10月20日　14:30 にお待ちしております'
const received = new Date('2026-10-17T00:00:00Z')

describe('validateBooking', () => {
  it('without an end time, the end is start + 60 minutes', () => {
    const result = validateBooking(base, received, body, 60)
    expect(result.ok && result.value.end).toBe('2026-10-20T15:30:00')
  })

  it('an explicit end time wins, falling back to the default if it is not after the start', () => {
    const explicit = validateBooking(
      { ...base, end_at: '2026-10-20T16:00' },
      received,
      body,
      60
    )
    expect(explicit.ok && explicit.value.end).toBe('2026-10-20T16:00:00')
    const invalid = validateBooking(
      { ...base, end_at: '2026-10-20T14:00' },
      received,
      body,
      60
    )
    expect(invalid.ok && invalid.value.end).toBe('2026-10-20T15:30:00')
  })

  it.each([null, 'tomorrow', '2026-02-31T10:00:00'])(
    'rejects an invalid start datetime: %s',
    (startAt) => {
      expect(
        validateBooking({ ...base, start_at: startAt }, received, body, 60).ok
      ).toBe(false)
    }
  )

  it('rejects an invalid time zone, the past and a too-distant future', () => {
    expect(
      validateBooking({ ...base, timezone: 'Mars/Base' }, received, body, 60).ok
    ).toBe(false)
    expect(
      validateBooking(
        { ...base, start_at: '2026-09-01T10:00:00' },
        received,
        body,
        60
      ).ok
    ).toBe(false)
    expect(
      validateBooking(
        { ...base, start_at: '2031-01-01T10:00:00' },
        received,
        body,
        60
      ).ok
    ).toBe(false)
  })

  it('rejects evidence absent from the body or containing no digits', () => {
    expect(
      validateBooking(
        { ...base, evidence: ['12月1日 9:00'] },
        received,
        body,
        60
      ).ok
    ).toBe(false)
    expect(
      validateBooking(
        { ...base, evidence: ['お待ちしております'] },
        received,
        body,
        60
      ).ok
    ).toBe(false)
  })

  it('rejects evidence consisting only of a relative date', () => {
    const relative = { ...base, evidence: ['来週火曜日 11時'] }
    expect(
      validateBooking(relative, received, 'ご希望: 来週火曜日 11時', 60).ok
    ).toBe(false)
  })

  it('rejects text that is too long for the calendar', () => {
    const long = 'あ'.repeat(201)
    expect(
      validateBooking({ ...base, subject: long }, received, body, 60)
    ).toEqual({ ok: false, reason: 'subject is longer than 200 characters' })
  })

  it('a cancel is not rejected for long text', () => {
    const long = 'あ'.repeat(201)
    const result = validateBooking(
      { ...base, subject: long },
      received,
      body,
      60,
      false
    )
    expect(result.ok).toBe(true)
  })
})
