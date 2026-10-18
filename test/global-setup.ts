import { spawn, type ChildProcess } from 'node:child_process'
import { mkdtempSync } from 'node:fs'
import net from 'node:net'
import os from 'node:os'
import path from 'node:path'

/** Azurite's fixed development account key. */
const DEV_KEY =
  'Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw=='

/**
 * Finds a free local TCP port.
 */
async function freePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = net.createServer()
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address() as net.AddressInfo
      server.close(() => resolve(port))
    })
  })
}

/**
 * Waits until the port accepts connections.
 */
async function waitForPort(port: number): Promise<void> {
  for (let attempt = 0; attempt < 100; attempt++) {
    const open = await new Promise<boolean>((resolve) => {
      const socket = net.connect(port, '127.0.0.1')
      socket.once('connect', () => {
        socket.destroy()
        resolve(true)
      })
      socket.once('error', () => resolve(false))
    })
    if (open) return
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
  throw new Error(`Azurite table service did not start on port ${port}`)
}

/**
 * Starts a throwaway Azurite table service so tests need no external setup.
 */
export default async function globalSetup(): Promise<void> {
  const port = await freePort()
  const location = mkdtempSync(path.join(os.tmpdir(), 'azurite-'))
  const child: ChildProcess = spawn(
    path.resolve('node_modules/.bin/azurite-table'),
    [
      '--silent',
      '--tableHost',
      '127.0.0.1',
      '--tablePort',
      String(port),
      '--location',
      location,
    ],
    { stdio: 'ignore' }
  )
  ;(globalThis as { __AZURITE__?: ChildProcess }).__AZURITE__ = child
  await waitForPort(port)
  process.env.TEST_STATE_CONNECTION = `DefaultEndpointsProtocol=http;AccountName=devstoreaccount1;AccountKey=${DEV_KEY};TableEndpoint=http://127.0.0.1:${port}/devstoreaccount1;`
}
