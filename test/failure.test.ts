import { classifyFailure, describeError } from '../src/core/failure'
import { FoundryError, FoundryOutputError } from '../src/integrations/foundry'
import { GraphError } from '../src/integrations/graph'

describe('classifyFailure', () => {
  it.each([400, 413, 422])('Foundry %i is per-mail', (status) => {
    expect(classifyFailure(new FoundryError(status))).toBe('per-mail')
  })

  it.each([401, 403, 404, 408, 409, 429, 500, 503])(
    'Foundry %i is treated as a configuration or outage problem',
    (status) => {
      expect(classifyFailure(new FoundryError(status))).toBe('systemic')
    }
  )

  it('a malformed model output is per-mail', () => {
    expect(classifyFailure(new FoundryOutputError('x'))).toBe('per-mail')
  })

  it.each([400, 401, 403, 404, 429, 503])(
    'Graph %i is treated as a systemic problem',
    (status) => {
      expect(classifyFailure(new GraphError(status, 'x'))).toBe('systemic')
    }
  )

  it('an unknown error is treated as systemic', () => {
    expect(classifyFailure(new Error('boom'))).toBe('systemic')
    expect(classifyFailure('string')).toBe('systemic')
  })
})

describe('describeError', () => {
  it('appends the error code from cause', () => {
    const error = new TypeError('fetch failed', {
      cause: { code: 'ENOTFOUND' },
    })
    expect(describeError(error)).toBe('fetch failed (ENOTFOUND)')
  })

  it('returns a string even for non-Error values', () => {
    expect(describeError(undefined)).toBe('non-Error value thrown')
  })
})
