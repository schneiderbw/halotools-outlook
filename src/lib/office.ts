// Promise wrappers + helpers around Office.context.mailbox.item.
//
// Read-surface only for v1 (mailRead context). Compose surface would need separate helpers.

export interface EmailContext {
  senderEmail: string;
  senderName: string;
  subject: string;
  conversationId: string;
  internetMessageId: string;
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
export function getCurrentEmailContext(): EmailContext | undefined {
  const item = Office.context.mailbox.item;
  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) return undefined;

  const from = (item as Office.MessageRead).from;
  if (!from) return undefined;

  return {
    senderEmail: from.emailAddress,
    senderName: from.displayName,
    subject: item.subject ?? "",
    conversationId: (item as Office.MessageRead).conversationId,
    internetMessageId: (item as Office.MessageRead).internetMessageId,
    itemId: item.itemId,
    receivedAt: (item as Office.MessageRead).dateTimeCreated,
  };
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

// ---------- Compose-surface helpers ----------
//
// These only work when the task pane is mounted under a mailCompose runtime.
// Each is guarded so a misuse from the read surface fails fast with a clear error
// rather than triggering an opaque Office.js exception.

/** True when the current item exposes the compose-mode recipient fields. */
function isComposeItem(): boolean {
  const item = Office.context?.mailbox?.item as Office.MessageCompose | undefined;
  return !!item && typeof (item as Office.MessageCompose).to?.getAsync === "function";
}

function getAsyncRecipients(
  field: Office.Recipients,
): Promise<Office.EmailAddressDetails[]> {
  return new Promise((resolve, reject) => {
    field.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error?.message ?? "Failed to read recipients"));
    });
  });
}

/**
 * Read To / Cc / Bcc on the current compose item.
 * Returns lowercased email addresses (display names stripped) for stable downstream lookup.
 */
export async function getRecipients(): Promise<{
  to: string[];
  cc: string[];
  bcc: string[];
}> {
  if (!isComposeItem()) {
    throw new Error("getRecipients is only available in the compose surface");
  }
  const item = Office.context.mailbox.item as Office.MessageCompose;
  const [to, cc, bcc] = await Promise.all([
    getAsyncRecipients(item.to),
    getAsyncRecipients(item.cc),
    getAsyncRecipients(item.bcc),
  ]);
  const flat = (arr: Office.EmailAddressDetails[]) =>
    arr.map((r) => r.emailAddress).filter(Boolean);
  return { to: flat(to), cc: flat(cc), bcc: flat(bcc) };
}

/** Insert an HTML fragment at the cursor in the compose body. */
export function insertIntoBody(html: string): Promise<void> {
  if (!isComposeItem()) {
    return Promise.reject(
      new Error("insertIntoBody is only available in the compose surface"),
    );
  }
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item?.body.setSelectedDataAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message ?? "Failed to insert content"));
      },
    );
  });
}

/** Save the in-progress compose draft and return its server-side itemId. */
export function saveDraft(): Promise<string> {
  if (!isComposeItem()) {
    return Promise.reject(new Error("saveDraft is only available in the compose surface"));
  }
  return new Promise((resolve, reject) => {
    (Office.context.mailbox.item as Office.MessageCompose).saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error?.message ?? "Failed to save draft"));
    });
  });
}

/** Read the current compose body. */
export function getComposeBody(format: "text" | "html" = "html"): Promise<string> {
  if (!isComposeItem()) {
    return Promise.reject(
      new Error("getComposeBody is only available in the compose surface"),
    );
  }
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

/** Read the compose subject (best-effort; returns empty string if unavailable). */
export function getComposeSubject(): Promise<string> {
  if (!isComposeItem()) return Promise.resolve("");
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item as Office.MessageCompose;
    item.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value ?? "");
      else resolve("");
    });
  });
}
