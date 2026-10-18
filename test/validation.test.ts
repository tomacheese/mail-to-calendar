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
  it('終了時刻が無い場合は開始 + 60 分になる', () => {
    const result = validateBooking(base, received, body, 60)
    expect(result.ok && result.value.end).toBe('2026-10-20T15:30:00')
  })

  it('明示された終了時刻を優先し、開始以前なら既定値に戻す', () => {
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
    '開始日時が不正な場合は拒否する: %s',
    (startAt) => {
      expect(
        validateBooking({ ...base, start_at: startAt }, received, body, 60).ok
      ).toBe(false)
    }
  )

  it('不正なタイムゾーン・過去・遠すぎる未来を拒否する', () => {
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

  it('本文に存在しない根拠・数字を含まない根拠を拒否する', () => {
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

  it('相対日付のみの根拠を拒否する', () => {
    const relative = { ...base, evidence: ['来週火曜日 11時'] }
    expect(
      validateBooking(relative, received, 'ご希望: 来週火曜日 11時', 60).ok
    ).toBe(false)
  })
})
