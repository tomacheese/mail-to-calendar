import { classifyFailure, describeError } from '../src/core/failure'
import { FoundryError, FoundryOutputError } from '../src/integrations/foundry'
import { GraphError } from '../src/integrations/graph'

describe('classifyFailure', () => {
  it.each([400, 413, 422])('Foundry %i はメール固有', (status) => {
    expect(classifyFailure(new FoundryError(status))).toBe('per-mail')
  })

  it.each([401, 403, 404, 408, 409, 429, 500, 503])(
    'Foundry %i は設定・障害として扱う',
    (status) => {
      expect(classifyFailure(new FoundryError(status))).toBe('systemic')
    }
  )

  it('モデル出力の不備はメール固有', () => {
    expect(classifyFailure(new FoundryOutputError('x'))).toBe('per-mail')
  })

  it.each([
    [400, 'per-mail'],
    [404, 'per-mail'],
    [401, 'systemic'],
    [403, 'systemic'],
    [429, 'systemic'],
    [503, 'systemic'],
  ])('Graph %i は %s', (status, kind) => {
    expect(classifyFailure(new GraphError(status, 'x'))).toBe(kind)
  })

  it('不明なエラーはシステム全体の問題として扱う', () => {
    expect(classifyFailure(new Error('boom'))).toBe('systemic')
    expect(classifyFailure('string')).toBe('systemic')
  })
})

describe('describeError', () => {
  it('cause のエラーコードを添える', () => {
    const error = new TypeError('fetch failed', {
      cause: { code: 'ENOTFOUND' },
    })
    expect(describeError(error)).toBe('fetch failed (ENOTFOUND)')
  })

  it('Error 以外でも文字列を返す', () => {
    expect(describeError(undefined)).toBe('non-Error value thrown')
  })
})
