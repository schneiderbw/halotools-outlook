// Promise wrappers + helpers around Office.context.mailbox.item.
//
// Read-surface only for v1 (mailRead context). Compose surface would need separate helpers.

export interface EmailContext {
  senderEmail: string;
  senderName: string;
  subject: string;
  conversationId: string;
  /** RFC 5322 Message-ID of the current message, with angle brackets stripped. */
  internetMessageId: string;
  /** Parent's Message-ID from the In-Reply-To header, angle brackets stripped. */
  inReplyTo?: string;
  /** Ancestor Message-IDs from the References header, angle brackets stripped. */
  references: string[];
  itemId: string;
  receivedAt: Date;
}

export interface FetchedAttachment {
  /** Filename (sanitized for Halo) */
  filename: string;
  contentType: string;
  /** Base64-encoded file content */
  base64: string;
  size: number;
  isInline: boolean;
}

export function isOfficeReady(): boolean {
  return typeof Office !== "undefined" && !!Office.context?.mailbox?.item;
}

export function awaitOffice(): Promise<void> {
  return new Promise((resolve) => {
    Office.onReady(() => resolve());
  });
}

/** Pull metadata from the currently selected message in the read pane. */
export async function getCurrentEmailContext(): Promise<EmailContext | undefined> {
  const item = Office.context.mailbox.item;
  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) return undefined;

  const msg = item as Office.MessageRead;
  const from = msg.from;
  if (!from) return undefined;

  const headers = await getInReplyToAndReferences(msg);

  return {
    senderEmail: from.emailAddress,
    senderName: from.displayName,
    subject: item.subject ?? "",
    conversationId: msg.conversationId,
    internetMessageId: stripAngleBrackets(msg.internetMessageId),
    inReplyTo: headers.inReplyTo,
    references: headers.references,
    itemId: item.itemId,
    receivedAt: msg.dateTimeCreated,
  };
}

/**
 * Read In-Reply-To and References headers from the current message.
 * Requires Mailbox 1.8+ (getInternetHeadersAsync). On older versions or any
 * failure we resolve to empty so callers can degrade gracefully.
 */
function getInReplyToAndReferences(
  msg: Office.MessageRead,
): Promise<{ inReplyTo?: string; references: string[] }> {
  return new Promise((resolve) => {
    const getter = (
      msg as unknown as {
        getAllInternetHeadersAsync?: (
          cb: (r: Office.AsyncResult<string>) => void,
        ) => void;
      }
    ).getAllInternetHeadersAsync;
    if (typeof getter !== "function") {
      resolve({ references: [] });
      return;
    }
    try {
      getter.call(msg, (result: Office.AsyncResult<string>) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          resolve({ references: [] });
          return;
        }
        resolve(parseThreadingHeaders(result.value));
      });
    } catch {
      resolve({ references: [] });
    }
  });
}

function parseThreadingHeaders(raw: string): { inReplyTo?: string; references: string[] } {
  // RFC 5322 headers come as a single CRLF-joined blob. Folded values (lines
  // starting with whitespace) belong to the previous header — unfold first.
  const unfolded = raw.replace(/\r\n[ \t]+/g, " ");
  const lines = unfolded.split(/\r?\n/);
  let inReplyToLine = "";
  let referencesLine = "";
  for (const line of lines) {
    const colon = line.indexOf(":");
    if (colon < 0) continue;
    const name = line.slice(0, colon).trim().toLowerCase();
    const value = line.slice(colon + 1).trim();
    if (name === "in-reply-to" && !inReplyToLine) inReplyToLine = value;
    else if (name === "references" && !referencesLine) referencesLine = value;
  }
  const inReplyToIds = extractMessageIds(inReplyToLine);
  const referenceIds = extractMessageIds(referencesLine);
  return {
    inReplyTo: inReplyToIds[0],
    references: referenceIds,
  };
}

function extractMessageIds(value: string): string[] {
  if (!value) return [];
  // Pull out anything that looks like a Message-ID. Most clients format these
  // as <local@domain>; some legacy senders omit the angle brackets.
  const bracketed = value.match(/<[^<>\s]+>/g);
  if (bracketed && bracketed.length) {
    return bracketed.map(stripAngleBrackets).filter(Boolean);
  }
  return value
    .split(/\s+/)
    .map((t) => stripAngleBrackets(t))
    .filter(Boolean);
}

function stripAngleBrackets(id: string | undefined): string {
  if (!id) return "";
  return id.trim().replace(/^<|>$/g, "");
}

/** Get the message body as plain text (default) or HTML. Async. */
export function getBody(format: "text" | "html" = "text"): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item?.body.getAsync(
      format === "html" ? Office.CoercionType.Html : Office.CoercionType.Text,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error?.message ?? "Failed to read body"));
      },
    );
  });
}

/** List attachments on the current message (metadata only). */
export function listAttachments(): Office.AttachmentDetails[] {
  const item = Office.context.mailbox.item as Office.MessageRead | undefined;
  return item?.attachments ?? [];
}

/** Fetch the content of a single attachment as base64. */
function fetchAttachmentContent(att: Office.AttachmentDetails): Promise<FetchedAttachment> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item?.getAttachmentContentAsync(
      att.id,
      { asyncContext: att },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error?.message ?? `Failed to read attachment ${att.name}`));
          return;
        }
        const c = result.value;
        // Office returns one of: Base64, Eml, ICalendar, Url.
        // We only handle Base64 for v1 — others surface as a per-attachment error.
        if (c.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
          reject(
            new Error(
              `Attachment "${att.name}" is ${c.format} — only base64 attachments are supported in v1.`,
            ),
          );
          return;
        }
        resolve({
          filename: sanitizeFilename(att.name),
          contentType: att.contentType,
          base64: c.content,
          size: att.size,
          isInline: att.isInline,
        });
      },
    );
  });
}

/**
 * Fetch all attachments as base64.
 * Inline images (signatures, embedded screenshots) are skipped by default to avoid noise.
 * Returns partial results on individual failures.
 */
export async function fetchAllAttachments(
  includeInline = false,
): Promise<{ attachments: FetchedAttachment[]; errors: string[] }> {
  const meta = listAttachments().filter((a) => includeInline || !a.isInline);
  const attachments: FetchedAttachment[] = [];
  const errors: string[] = [];
  for (const m of meta) {
    try {
      attachments.push(await fetchAttachmentContent(m));
    } catch (e) {
      errors.push((e as Error).message);
    }
  }
  return { attachments, errors };
}

/** Get the email address of the current Outlook user. */
export function getCurrentUserEmail(): string | undefined {
  return Office.context.mailbox.userProfile?.emailAddress;
}

/** Extract the domain part from an email address (lowercased). */
export function domainOf(email: string): string {
  const at = email.lastIndexOf("@");
  return at >= 0 ? email.slice(at + 1).toLowerCase() : "";
}

// ---------- helpers ----------

function sanitizeFilename(name: string): string {
  return name
    .replace(/[\/\\?<>:*|"]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 200);
}
