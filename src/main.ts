import { readFileSync } from 'node:fs'
import { loadConfig } from './config'
import { processMail, runScan } from './core/process-mail'
import { Foundry } from './integrations/foundry'
import { Graph } from './integrations/graph'
import { State } from './integrations/state'

/** Subject keywords used by the sample command to pick booking-looking mail. */
const BOOKING_SUBJECT =
  /予約|ご予約|予定|来院|診察|受診|リマインド|reservation|booking|appointment|itinerary|チェックイン|面談|ご来店/i

/**
 * Connectivity check: Graph token and Inbox count, plus Foundry analysis of a synthetic mail.
 * Never reads real mail bodies.
 */
async function check(): Promise<void> {
  const config = loadConfig()
  console.log('inbox items:', await new Graph(config).inboxCount())
  const analysis = await new Foundry(config).analyze({
    subject: '【ご予約確認】テスト歯科',
    from: 'noreply@example.com',
    receivedAt: new Date(),
    bodyText:
      'ご予約が確定しました。\n日時: 2099年1月5日(月) 14:30\n予約番号: T-0001\n場所: 東京都千代田区1-1',
  })
  console.log('foundry:', analysis.action, analysis.bookings[0]?.start_at)
}

/**
 * Analyses a local text file with Foundry and prints the raw result.
 *
 * @param file - Path of the text file to analyse
 */
async function analyze(file: string): Promise<void> {
  const result = await new Foundry(loadConfig()).analyze({
    subject: '',
    from: '',
    receivedAt: new Date(),
    bodyText: readFileSync(file, 'utf8'),
  })
  console.log(JSON.stringify(result, null, 2))
}

/**
 * Analyses existing booking-looking mail. Forced to dry-run unless --write is given.
 *
 * @param limit - Maximum number of mails per folder
 * @param write - When true, CALENDAR_WRITE_MODE applies instead of dry-run
 */
async function sample(limit: number, write: boolean): Promise<void> {
  const loaded = loadConfig()
  const config = write
    ? loaded
    : { ...loaded, CALENDAR_WRITE_MODE: 'dry-run' as const }
  const graph = new Graph(config)
  const foundry = new Foundry(config)
  const state = new State(config)
  await state.init()

  const picks = []
  for (const folder of ['inbox', 'deleteditems']) {
    const messages = await graph.listFolder(folder, 150)
    const hits = messages.filter((m) => BOOKING_SUBJECT.test(m.subject))
    picks.push(...hits.slice(0, limit).map((m) => ({ ...m, folder })))
  }
  console.log(`candidates: ${picks.length}`)
  for (const pick of picks) {
    try {
      const mail = await graph.getMail(pick.id)
      const analysis = await foundry.analyze(mail)
      const ops = await processMail(config, graph, state, mail, analysis)
      console.log(
        `- [${pick.folder}] ${pick.receivedDateTime.slice(0, 10)} "${pick.subject.slice(0, 50)}" → ${analysis.action} | ${analysis.reason.slice(0, 80)}`
      )
      for (const op of ops) {
        console.log(
          `    ${op.op} ${op.start ?? ''} ${(op.subject ?? '').slice(0, 40)} ${op.note ?? ''}`
        )
      }
    } catch (error: unknown) {
      console.log(
        `- [${pick.folder}] "${pick.subject.slice(0, 40)}" failed: ${(error as Error).message}`
      )
    }
  }
  console.log('foundry calls:', foundry.calls)
}

/**
 * Lists the plans remembered by dry-run runs. Contains booking details, so for local use only.
 */
async function report(): Promise<void> {
  const state = new State(loadConfig())
  await state.init()
  const rows = await state.listDryRunPlans()
  for (const row of rows) {
    console.log(row.receivedAt.slice(0, 16), row.result, row.plan)
  }
}

/**
 * Processes the Inbox delta since the last run.
 */
async function scan(): Promise<void> {
  const summary = await runScan(loadConfig())
  console.log(JSON.stringify(summary, null, 2))
  process.exitCode = summary.failures > 0 ? 1 : 0
}

/**
 * Entry point.
 */
async function main(): Promise<void> {
  const [command = 'scan', argument] = process.argv.slice(2)
  switch (command) {
    case 'check': {
      await check()
      break
    }
    case 'analyze': {
      if (!argument) throw new Error('usage: analyze <file>')
      await analyze(argument)
      break
    }
    case 'sample': {
      const limit = Number(
        argument && !argument.startsWith('--') ? argument : 8
      )
      await sample(limit, process.argv.includes('--write'))
      break
    }
    case 'report': {
      await report()
      break
    }
    case 'scan': {
      await scan()
      break
    }
    default: {
      throw new Error(
        'usage: main.ts [check | analyze <file> | sample [n] [--write] | report | scan]'
      )
    }
  }
}

main().catch((error: unknown) => {
  console.error(error)
  // eslint-disable-next-line unicorn/no-process-exit
  process.exit(1)
})
