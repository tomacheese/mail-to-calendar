import type { Config } from '../config'

const BASE = 'https://graph.microsoft.com/v1.0'

/** Reference to a mail returned by the delta query. */
export interface MailRef {
  id: string
  receivedDateTime?: string
  isDraft?: boolean
  removed: boolean
}
/** Mail content fetched for analysis. */
export interface Mail {
  subject: string
  from: string
  receivedAt: Date
  bodyText: string
}
/** Minimal event shape returned by create/update calls. */
export interface GraphEvent {
  id: string
  lastModifiedDateTime: string
  subject: string
}
/** Event fields we manage, as currently held by Outlook. */
export interface EventView {
  id: string
  subject: string
  showAs: string
  bodyPreview: string
  start: { dateTime: string }
  end: { dateTime: string }
  location?: { displayName?: string }
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/** Graph HTTP failure carrying status and error code only (bodies may contain mail content). */
export class GraphError extends Error {
  constructor(
    public status: number,
    msg: string
  ) {
    super(msg)
  }
}

/** Microsoft Graph client using client-credentials auth against one mailbox. */
export class Graph {
  private tok?: { v: string; exp: number }
  private user: string
  constructor(private cfg: Config) {
    this.user = `${BASE}/users/${encodeURIComponent(cfg.GRAPH_MAILBOX)}`
  }

  /** Returns a cached app-only access token, refreshing it shortly before expiry. */
  private async token(): Promise<string> {
    if (this.tok && this.tok.exp > Date.now() + 60_000) return this.tok.v
    const res = await fetch(
      `https://login.microsoftonline.com/${this.cfg.GRAPH_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        body: new URLSearchParams({
          client_id: this.cfg.GRAPH_CLIENT_ID,
          client_secret: this.cfg.GRAPH_CLIENT_SECRET,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials',
        }),
      }
    )
    const j = (await res.json()) as {
      access_token?: string
      expires_in?: number
      error?: string
      error_description?: string
    }
    if (!j.access_token)
      throw new Error(
        `Graph token failed: ${j.error} ${j.error_description?.split('\r\n', 1)[0]}`
      )
    this.tok = {
      v: j.access_token,
      exp: Date.now() + (j.expires_in ?? 3600) * 1000,
    }
    return j.access_token
  }

  /** Graph call with 429/5xx retry (honours Retry-After). `url` may be relative to the mailbox or absolute. */
  async req<T>(
    method: string,
    url: string,
    body?: unknown,
    prefer?: string
  ): Promise<T> {
    const full = url.startsWith('http') ? url : this.user + url
    for (let attempt = 0; ; attempt++) {
      const res = await fetch(full, {
        method,
        headers: {
          Authorization: `Bearer ${await this.token()}`,
          'Content-Type': 'application/json',
          ...(prefer && { Prefer: prefer }),
        },
        body: body === undefined ? undefined : JSON.stringify(body),
      })
      if ([429, 502, 503, 504].includes(res.status) && attempt < 4) {
        await sleep(
          Math.max(
            Number(res.headers.get('retry-after')) * 1000 || 0,
            1000 * 2 ** attempt
          )
        )
        continue
      }
      if (!res.ok) {
        // Status + Graph error code only: bodies may contain mail content.
        const code = (
          (await res.json().catch(() => ({}))) as { error?: { code?: string } }
        ).error?.code
        throw new GraphError(
          res.status,
          `Graph ${method} ${res.status} ${code ?? ''}`
        )
      }
      return (res.status === 204 ? undefined : await res.json()) as T
    }
  }

  /** Returns the number of items in the Inbox. */
  async inboxCount(): Promise<number> {
    const folder = await this.req<{ totalItemCount: number }>(
      'GET',
      '/mailFolders/inbox?$select=totalItemCount'
    )
    return folder.totalItemCount
  }

  /** Follows nextLink to the end. Pass the saved deltaLink for an incremental run, nothing for a baseline walk. */
  async inboxDelta(
    deltaLink = '/mailFolders/inbox/messages/delta?$select=receivedDateTime,isDraft'
  ): Promise<{ messages: MailRef[]; deltaLink: string }> {
    let url = deltaLink
    const messages: MailRef[] = []
    for (;;) {
      const page = await this.req<{
        value: Record<string, unknown>[]
        '@odata.nextLink'?: string
        '@odata.deltaLink'?: string
      }>('GET', url, undefined, 'IdType="ImmutableId", odata.maxpagesize=50')
      for (const m of page.value) {
        messages.push({
          id: m.id as string,
          receivedDateTime: m.receivedDateTime as string | undefined,
          isDraft: m.isDraft as boolean | undefined,
          removed: '@removed' in m,
        })
      }
      const next = page['@odata.nextLink']
      if (next) {
        url = next
        continue
      }
      const done = page['@odata.deltaLink']
      if (!done) {
        throw new Error('Graph delta page has neither nextLink nor deltaLink')
      }
      return { messages, deltaLink: done }
    }
  }

  /** Newest `top` messages of a well-known folder (inbox, deleteditems), metadata only. */
  async listFolder(
    folder: string,
    top: number
  ): Promise<{ id: string; subject: string; receivedDateTime: string }[]> {
    const r = await this.req<{
      value: { id: string; subject: string; receivedDateTime: string }[]
    }>(
      'GET',
      `/mailFolders/${folder}/messages?$top=${top}&$select=subject,receivedDateTime&$orderby=receivedDateTime desc`,
      undefined,
      'IdType="ImmutableId"'
    )
    return r.value
  }

  /** Fetches subject, sender, receipt time and plain-text body of a mail. */
  async getMail(id: string): Promise<Mail> {
    const m = await this.req<{
      subject: string | null
      receivedDateTime: string
      from?: { emailAddress?: { address?: string } }
      body: { content: string }
    }>(
      'GET',
      `/messages/${encodeURIComponent(id)}?$select=subject,from,receivedDateTime,body`,
      undefined,
      'IdType="ImmutableId", outlook.body-content-type="text"'
    )
    return {
      subject: m.subject ?? '',
      from: m.from?.emailAddress?.address ?? '',
      receivedAt: new Date(m.receivedDateTime),
      bodyText: m.body.content.slice(0, 20_000),
    }
  }

  /** Creates an event in the main calendar; transactionId makes retries idempotent. */
  createEvent(ev: object, transactionId: string) {
    return this.req<GraphEvent>('POST', '/calendar/events', {
      ...ev,
      transactionId,
    })
  }
  /** Fields we manage, with times rendered in `tz` so they compare equal to what we wrote. */
  getEvent(id: string, tz: string) {
    return this.req<EventView>(
      'GET',
      `/events/${encodeURIComponent(id)}?$select=subject,start,end,location,showAs,bodyPreview`,
      undefined,
      `outlook.timezone="${tz}"`
    )
  }
  /** Moves the event to Deleted Items (recoverable from Outlook). */
  async deleteEvent(id: string): Promise<void> {
    await this.req<undefined>('DELETE', `/events/${encodeURIComponent(id)}`)
  }
  /** Updates the given fields of an event. */
  patchEvent(id: string, ev: object) {
    return this.req<GraphEvent>(
      'PATCH',
      `/events/${encodeURIComponent(id)}`,
      ev
    )
  }
}
