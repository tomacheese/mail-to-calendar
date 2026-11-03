# AGENTS.md

## Language

- Write all deliverables in English: commit messages, PR/Issue text, review comments, code comments and JSDoc, and docs (README, AGENTS.md, infra/README.md).
- Keep as-is: Japanese data the code matches against or emits (the mail-parsing regex `BOOKING_SUBJECT` in `src/main.ts`, the `【仮】` event-title prefix and the `予約自動登録` calendar category in `src/core/process-mail.ts`), Japanese test fixtures (mail bodies, provider names), quotations, technical identifiers and generated files.
- Error and log messages: English.

## Overview

Azure Functions (v4, Node.js, TypeScript) app that reads booking mail from an Exchange Online Inbox (Microsoft Graph delta), analyzes it with Microsoft Foundry, and creates, updates or deletes events in the Outlook main calendar. State lives in Azure Table Storage. See `README.md` for behavior and environment variables.

## Development commands

pnpm only (`.node-version` pins Node.js).

- `pnpm install`
- `pnpm lint` (prettier + eslint + `tsc --noEmit`; lints `src/` only), `pnpm fix`
- `pnpm test` (jest; `test/global-setup.ts` starts Azurite automatically)
- `pnpm build` (`tsc -p tsconfig.build.json` into `dist/`)
- `pnpm start <check|scan|analyze <file>|sample [n] [--write]|report>` (local CLI, reads `.env`)
- `pnpm test:e2e` calls the real Foundry and incurs cost; do not run without approval

## Architecture

- `src/main.ts`: local CLI. `src/functions/timer.ts`: 5-minute timer that calls `runScan` only when `MAIL_POLLING_ENABLED=true`.
- `src/config.ts`: zod-validated environment; unset `CALENDAR_WRITE_MODE` falls back to `dry-run`.
- `src/core/`: `process-mail.ts` (scan loop, booking identity by normalized provider + reservation number, pending ops), `validation.ts` (re-validates model output), `failure.ts` (per-mail vs systemic failures).
- `src/integrations/`: `graph.ts`, `foundry.ts`, `state.ts` (tables: SyncState, ProcessedMessages, Bookings, PendingOps).
- `test/`: jest suites using in-memory Graph/Foundry stubs plus Azurite; `test/corpus.ts` is the manual E2E.
- `infra/`: Bicep; `.github/workflows/`: CI (shared template) and Azure Functions deploy.

## Conventions

- TypeScript, no semicolons (prettier); document exported/top-level functions with JSDoc in English.
- No behavior change in documentation-only work; keep comments for design intent and constraints, not for restating code.
- Commits follow Conventional Commits.

## Safety rules

- Never commit `.env`, secrets, mail content or event details; never log credential values. In the Azure Functions timer path and any production log, also never log mail content or event details (logs carry counts, hashed message ids and sanitized errors). The local CLI (`sample`, `scan`, `report`) prints truncated subjects and booking details to the console for local use only; do not paste that output into PRs, Issues or shared logs.
- Do not change `CALENDAR_WRITE_MODE` of a real environment away from `dry-run`, and do not run `scan`/`sample --write` against a real mailbox with `create-only`/`full`, without explicit approval.
- Do not touch the production calendar or mailbox; do not hand-edit events the app manages (the app deliberately leaves hand-edited events alone).
