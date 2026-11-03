# mail-to-calendar

Analyzes booking and appointment emails that arrive in an Exchange Online Inbox with Microsoft Foundry and automatically registers them as events in the Outlook main calendar. It runs as a 5-minute Timer Trigger on Azure Functions (v4, Node.js, Flex Consumption).

- Mail retrieval: Microsoft Graph Inbox Delta (new mail only; the first run just establishes a baseline and does not register past mail)
- Analysis: classification and date/time extraction with an existing Microsoft Foundry deployment (API key authentication); the app then re-validates the date/time, the evidence and booking identity
- Registration: confirmed bookings are `busy`; tentative ones get a `【仮】` subject prefix and `tentative`. Updates modify the event, cancellations delete it (moved to Deleted Items)
- A booking is identified by "provider + reservation number". If another provider's booking has the same reservation number, the existing event is not updated or deleted and the mail is held (even the same provider is treated as a different one if its name differs substantially)
- Events edited by hand are never overwritten; only events created automatically are managed

## Setup

### Requirements

- Node.js (see `.node-version` for the version)
- pnpm (enforced by `only-allow pnpm` in `preinstall`)

```bash
pnpm install
```

### Environment variables

Locally, set them in `.env` (not tracked by Git); for Functions, use `local.settings.json` `Values` or app settings. The template is [`.env.example`](./.env.example).

| Variable                                                      | Required | Description                                                                   |
| ------------------------------------------------------------- | -------- | ----------------------------------------------------------------------------- |
| `GRAPH_TENANT_ID` / `GRAPH_CLIENT_ID` / `GRAPH_CLIENT_SECRET` | yes      | Graph app registration (Client ID is a GUID, Secret is the secret value)      |
| `GRAPH_MAILBOX`                                               | yes      | UPN of the target mailbox                                                     |
| `FOUNDRY_ENDPOINT` / `FOUNDRY_DEPLOYMENT` / `FOUNDRY_API_KEY` | yes      | Connection details of the existing Foundry resource                           |
| `STATE_STORAGE_CONNECTION`                                    | yes      | State storage (Azurite locally, the app's storage account in production)      |
| `CALENDAR_WRITE_MODE`                                         | -        | `dry-run` (default) / `create-only` / `full`                                  |
| `MAIL_POLLING_ENABLED`                                        | -        | Automatic polling by the timer (default `false`)                              |
| `PROCESS_EXISTING_MESSAGES`                                   | -        | Whether to process existing mail on the first run (default `false`)           |
| `DEFAULT_EVENT_DURATION_MINUTES`                              | -        | Duration when the end time is unknown (default 60)                            |
| `FOUNDRY_MAX_CALLS_PER_RUN`                                   | -        | Maximum inference calls per run (app default 50, Bicep production default 20) |
| `STATE_TABLE_PREFIX`                                          | -        | Alphanumeric table-name prefix (to isolate state within one storage account)  |

### Development commands

```bash
pnpm start check        # Check Graph / Foundry connectivity (does not read real mail bodies)
pnpm start scan         # Process the Inbox delta (mode follows CALENDAR_WRITE_MODE)
pnpm start analyze <file>  # Analyze a local text file with Foundry and print the raw result
pnpm start sample 8     # Analyze existing mail in dry-run (add --write to apply CALENDAR_WRITE_MODE)
pnpm start report       # Show the plans recorded in dry-run
pnpm start:func         # Build and start the local Functions host
pnpm build              # tsc build (dist/functions/*.js)
pnpm test               # jest (starts Azurite automatically)
pnpm test:e2e           # E2E with synthetic mail against the real Foundry (Graph is stubbed, incurs cost; needs .env and Azurite)
pnpm lint               # prettier / eslint / tsc
pnpm fix                # prettier / eslint auto-fix
```

## Update and cancel mails with no matching booking

With `CALENDAR_WRITE_MODE=full`, an update or cancel mail that matches no managed booking does not touch the calendar; it is queued as a pending item in the `PendingOps` table of the state storage. Pending items are re-evaluated for 14 days at the end of every scan and applied once the booking is created.

- If the booking was already updated or created by a newer mail, the older update/cancel is discarded without being applied (`superseded`)
- An item is closed as `done` when applied, `expired` after 14 days, or `failed` after 3 failures caused by corrupted stored content or item-specific Graph errors. Closed items remain as a record
- Nothing is queued in `dry-run` / `create-only`

## Directory layout

```text
src/
  main.ts                # Local CLI (check / analyze / sample / report / scan)
  config.ts              # Environment variable validation
  functions/timer.ts     # 5-minute Timer Trigger (thin entry point calling the shared engine)
  core/                  # Booking processing engine, validation, state transitions
  integrations/          # Graph / Foundry / Azure Tables
test/                    # jest (Azurite is started automatically by global-setup)
infra/                   # Bicep (main.bicep / main.bicepparam / README.md)
```

## Deployment

See [`infra/README.md`](./infra/README.md) for provisioning the infrastructure and [`.github/workflows/azure-functions-deploy.yml`](./.github/workflows/azure-functions-deploy.yml) (push to `main` / `master`) for deploying the app. The initial production mode must be `dry-run`; switching to writing to the main calendar requires separate approval.

## License

[MIT](./LICENSE)
