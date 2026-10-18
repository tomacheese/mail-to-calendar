import { FoundryError, FoundryOutputError } from '../integrations/foundry'
import { GraphError } from '../integrations/graph'

/**
 * per-mail: the same mail fails the same way on retry, so it is retried a few times and then parked.
 * systemic: configuration or outage, so it must never be parked and must hold the cursor and alert.
 */
export type FailureKind = 'per-mail' | 'systemic'

const GRAPH_SYSTEMIC_4XX = new Set([401, 403, 408, 429])

/** Decides whether a failure belongs to one mail or to the whole system. */
export function classifyFailure(error: unknown): FailureKind {
  if (error instanceof FoundryOutputError) return 'per-mail'
  if (error instanceof FoundryError) {
    return error.permanent ? 'per-mail' : 'systemic'
  }
  if (error instanceof GraphError) {
    const clientError = error.status >= 400 && error.status < 500
    return clientError && !GRAPH_SYSTEMIC_4XX.has(error.status)
      ? 'per-mail'
      : 'systemic'
  }
  return 'systemic'
}

/** One-line description for logs: message plus the network error code hidden in `cause`. */
export function describeError(error: unknown): string {
  if (!(error instanceof Error)) return 'non-Error value thrown'
  const cause = error.cause as { code?: string } | undefined
  return cause?.code ? `${error.message} (${cause.code})` : error.message
}
