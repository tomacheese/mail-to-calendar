import { FoundryError, FoundryOutputError } from '../integrations/foundry'

/**
 * per-mail: tied to this one mail, so it is retried a few times and then parked (permanent Foundry rejections are parked at once).
 * systemic: configuration or outage, so it must never be parked and must hold the cursor and alert.
 */
export type FailureKind = 'per-mail' | 'systemic'

/** Decides whether a failure belongs to one mail or to the whole system. */
export function classifyFailure(error: unknown): FailureKind {
  if (error instanceof FoundryOutputError) return 'per-mail'
  if (error instanceof FoundryError) {
    return error.permanent ? 'per-mail' : 'systemic'
  }
  // Graph failures depend on the call, not on the mail, so they hold the cursor and alert. A mail that is gone (404/410) is handled in runScan.
  return 'systemic'
}

/** One-line description for logs: message plus the network error code hidden in `cause`. */
export function describeError(error: unknown): string {
  if (!(error instanceof Error)) return 'non-Error value thrown'
  const cause = error.cause as { code?: string } | undefined
  return cause?.code ? `${error.message} (${cause.code})` : error.message
}
