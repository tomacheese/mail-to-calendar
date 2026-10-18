import type { ChildProcess } from 'node:child_process'

/**
 * Stops the Azurite process started by global-setup.
 */
export default function globalTeardown(): void {
  ;(globalThis as { __AZURITE__?: ChildProcess }).__AZURITE__?.kill()
}
