import { z } from 'zod'

const bool = z
  .enum(['true', 'false'])
  .default('false')
  .transform((v) => v === 'true')

const schema = z.object({
  APP_ENV: z.string().default('local'),
  // Guard against swapped CLIENT_ID / CLIENT_SECRET values (the secret *ID* is also a GUID, so check the secret shape too).
  GRAPH_TENANT_ID: z.guid(),
  GRAPH_CLIENT_ID: z.guid({
    error: 'GRAPH_CLIENT_ID must be the Application (client) ID GUID',
  }),
  GRAPH_CLIENT_SECRET: z
    .string()
    .min(20)
    .refine((v) => !z.guid().safeParse(v).success, {
      error:
        'GRAPH_CLIENT_SECRET looks like a GUID (secret ID?) — use the secret *value*',
    }),
  GRAPH_MAILBOX: z.string().min(1),
  FOUNDRY_ENDPOINT: z.url(),
  FOUNDRY_DEPLOYMENT: z.string().min(1),
  FOUNDRY_API_KEY: z.string().min(1),
  STATE_STORAGE_CONNECTION: z.string().min(1),
  // Alphanumeric prefix for table names; isolates test/PoC state inside one storage account.
  STATE_TABLE_PREFIX: z
    .string()
    .regex(/^[A-Za-z0-9]*$/)
    .default(''),
  // Fail closed: unset = dry-run, anything unknown = startup error.
  CALENDAR_WRITE_MODE: z
    .enum(['dry-run', 'create-only', 'full'])
    .default('dry-run'),
  MAIL_POLLING_ENABLED: bool,
  PROCESS_EXISTING_MESSAGES: bool,
  DEFAULT_EVENT_DURATION_MINUTES: z.coerce
    .number()
    .int()
    .positive()
    .default(60),
  FOUNDRY_MAX_CALLS_PER_RUN: z.coerce.number().int().positive().default(50),
})

export type Config = z.infer<typeof schema>

/** Validates environment variables into a typed config. Throws with variable names only, never values. */
export function loadConfig(env: NodeJS.ProcessEnv = process.env): Config {
  const r = schema.safeParse(env)
  if (!r.success) {
    // Only variable names and messages — never values.
    throw new Error(
      'Invalid config:\n' +
        r.error.issues
          .map((i) => `  ${i.path.join('.')}: ${i.message}`)
          .join('\n')
    )
  }
  return r.data
}
