import { app } from '@azure/functions'
import { loadConfig } from '../config'
import { runScan } from '../core/process-mail'

app.timer('scanInbox', {
  schedule: '0 */5 * * * *',
  handler: async (_t, ctx) => {
    const cfg = loadConfig()
    if (!cfg.MAIL_POLLING_ENABLED) return
    // runScan logs counts, hashed message ids and error classes, never mail content or event details.
    const s = await runScan(cfg, (m) => {
      ctx.log(m)
    })
    ctx.log(
      `scan done: seen=${s.seen} processed=${s.processed} ops=${s.ops.length} applied=${s.ops.filter((o) => o.applied).length} failures=${s.failures} deferred=${s.deferred}`
    )
    if (s.failures) throw new Error('scan had failures; cursor not advanced')
  },
})
