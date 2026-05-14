import { useEffect, useState } from "react";
import {
  Button,
  makeStyles,
  tokens,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Field,
  Combobox,
  Option,
  Spinner,
  MessageBar,
  MessageBarBody,
  Text,
  Switch,
  Input,
} from "@fluentui/react-components";
import { Add24Regular, Attach24Regular } from "@fluentui/react-icons";
import {
  appendAction,
  createTicket,
  listTicketTypes,
  stampOutlookThreadFields,
} from "../lib/halo-api";
import {
  getBody,
  fetchAllAttachments,
  listAttachments,
  type EmailContext,
  type FetchedAttachment,
} from "../lib/office";
import type {
  HaloTicket,
  HaloUser,
  HaloClient,
  HaloTicketType,
  HaloAttachmentInline,
} from "../types/halo";
import { getDefaults, setDefaults } from "../lib/defaults";

const CUSTOM_FIELD_CONVERSATION_ID = "CFOutlookConversationId";
const CUSTOM_FIELD_MESSAGE_ID = "CFOutlookInternetMessageId";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  buttons: {
    display: "flex",
    gap: "8px",
  },
  buttonFull: {
    flex: 1,
  },
  successText: {
    color: tokens.colorPaletteGreenForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  toggleRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    paddingTop: "4px",
  },
  toggleLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  hint: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
});

interface Props {
  email: EmailContext;
  client?: HaloClient;
  contact?: HaloUser;
  candidateTickets: HaloTicket[];
  /** When true, Append is primary and Create is secondary (use when a thread match exists). */
  preferAppend?: boolean;
}

export function LogActions({
  email,
  client,
  contact,
  candidateTickets,
  preferAppend,
}: Props) {
  const styles = useStyles();
  const [success, setSuccess] = useState<string | undefined>();
  const [error, setError] = useState<string | undefined>();
  const [warning, setWarning] = useState<string | undefined>();

  const announce = (kind: "success" | "error" | "warning", msg: string) => {
    setSuccess(undefined);
    setError(undefined);
    setWarning(undefined);
    if (kind === "success") setSuccess(msg);
    else if (kind === "error") setError(msg);
    else setWarning(msg);
    if (kind === "success") setTimeout(() => setSuccess(undefined), 5000);
  };

  return (
    <div className={styles.root}>
      <div className={styles.buttons}>
        <AppendDialog
          email={email}
          tickets={candidateTickets}
          onResult={announce}
          triggerClass={styles.buttonFull}
          appearance={preferAppend ? "primary" : "secondary"}
        />
        <CreateDialog
          email={email}
          client={client}
          contact={contact}
          onResult={announce}
          triggerClass={styles.buttonFull}
          appearance={preferAppend ? "secondary" : "primary"}
          dedupWarning={preferAppend}
        />
      </div>

      {success && <Text className={styles.successText}>{success}</Text>}
      {warning && (
        <MessageBar intent="warning">
          <MessageBarBody>{warning}</MessageBarBody>
        </MessageBar>
      )}
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}

// ---------- Append to ticket ----------

function AppendDialog({
  email,
  tickets,
  onResult,
  triggerClass,
  appearance = "secondary",
}: {
  email: EmailContext;
  tickets: HaloTicket[];
  onResult: (kind: "success" | "error" | "warning", msg: string) => void;
  triggerClass: string;
  appearance?: "primary" | "secondary";
}) {
  const [selectedId, setSelectedId] = useState<number | undefined>();
  const [busy, setBusy] = useState(false);
  const [open, setOpen] = useState(false);
  const [internalNote, setInternalNote] = useState(false);
  const [includeAttachments, setIncludeAttachments] = useState(
    getDefaults().includeAttachmentsByDefault ?? true,
  );
  const [timeMinutes, setTimeMinutes] = useState<string>("");
  const attachmentCount = listAttachments().filter((a) => !a.isInline).length;

  const submit = async () => {
    if (!selectedId) return;
    setBusy(true);
    try {
      const html = await getBody("html");
      let attachments: HaloAttachmentInline[] = [];
      let attachWarning: string | undefined;
      if (includeAttachments && attachmentCount > 0) {
        const fetched = await fetchAllAttachments();
        attachments = fetched.attachments.map(toHaloAttachment);
        if (fetched.errors.length > 0) {
          attachWarning = `Some attachments couldn't be included: ${fetched.errors.join("; ")}`;
        }
      }

      // Convert minutes input to decimal hours for Halo's time_taken field
      const time = parseFloat(timeMinutes);
      const timeHours = Number.isFinite(time) && time > 0 ? time / 60 : undefined;

      const action = await appendAction({
        ticket_id: selectedId,
        outcome: getDefaults().defaultAppendOutcome ?? "Email Received",
        note: html,
        hiddenfromuser: internalNote,
        emailfrom: email.senderEmail,
        emailfromname: email.senderName,
        emailsubject: email.subject,
        attachments: attachments.length ? attachments : undefined,
        time_taken: timeHours,
      });

      // Stamp the ticket so future emails in this conversation auto-thread to it.
      // Non-fatal: if this fails, the append still succeeded.
      stampOutlookThreadFields(
        selectedId,
        email.conversationId,
        email.internetMessageId,
      ).catch(() => {
        /* swallow — append succeeded, threading is best-effort */
      });

      if (attachWarning) {
        onResult("warning", `Appended to #${action.ticket_id}, but: ${attachWarning}`);
      } else {
        onResult("success", `Appended to #${action.ticket_id}`);
      }
      setOpen(false);
    } catch (e) {
      onResult("error", (e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, d) => setOpen(d.open)}>
      <DialogTrigger disableButtonEnhancement>
        <Button
          appearance={appearance}
          icon={<Attach24Regular />}
          className={triggerClass}
          disabled={tickets.length === 0}
        >
          Append
        </Button>
      </DialogTrigger>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Append email to ticket</DialogTitle>
          <DialogContent>
            <Field label="Ticket" required>
              <Combobox
                placeholder="Select a ticket"
                onOptionSelect={(_, d) =>
                  setSelectedId(d.optionValue ? Number(d.optionValue) : undefined)
                }
              >
                {tickets.map((t) => (
                  <Option key={t.id} value={String(t.id)} text={`#${t.id} ${t.summary}`}>
                    #{t.id} · {t.summary}
                    {t.statusname ? ` · ${t.statusname}` : ""}
                  </Option>
                ))}
              </Combobox>
            </Field>

            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 12 }}>
              <Switch
                checked={internalNote}
                onChange={(_, d) => setInternalNote(d.checked)}
                label="Private note (hidden from customer)"
              />
              <Switch
                checked={includeAttachments}
                onChange={(_, d) => setIncludeAttachments(d.checked)}
                disabled={attachmentCount === 0}
                label={
                  attachmentCount === 0
                    ? "No attachments"
                    : `Include ${attachmentCount} attachment${attachmentCount === 1 ? "" : "s"}`
                }
              />
              <Field label="Time spent (minutes)" hint="Optional — recorded on the action">
                <Input
                  type="number"
                  value={timeMinutes}
                  onChange={(_, d) => setTimeMinutes(d.value)}
                  placeholder="e.g. 15"
                />
              </Field>
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => setOpen(false)} disabled={busy}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              onClick={submit}
              disabled={!selectedId || busy}
              icon={busy ? <Spinner size="tiny" /> : undefined}
            >
              {busy ? "Appending…" : "Append"}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

// ---------- Create new ticket ----------

function CreateDialog({
  email,
  client,
  contact,
  onResult,
  triggerClass,
  appearance = "primary",
  dedupWarning = false,
}: {
  email: EmailContext;
  client?: HaloClient;
  contact?: HaloUser;
  onResult: (kind: "success" | "error" | "warning", msg: string) => void;
  triggerClass: string;
  appearance?: "primary" | "secondary";
  dedupWarning?: boolean;
}) {
  const [busy, setBusy] = useState(false);
  const [open, setOpen] = useState(false);
  const [summary, setSummary] = useState(email.subject);
  const [ticketTypes, setTicketTypes] = useState<HaloTicketType[]>([]);
  const [ticketTypeId, setTicketTypeId] = useState<number | undefined>(
    getDefaults().defaultTicketTypeId,
  );
  const [includeAttachments, setIncludeAttachments] = useState(
    getDefaults().includeAttachmentsByDefault ?? true,
  );
  const [loadError, setLoadError] = useState<string | undefined>();
  const attachmentCount = listAttachments().filter((a) => !a.isInline).length;

  useEffect(() => {
    if (!open) return;
    setSummary(email.subject);
    listTicketTypes()
      .then((types) => {
        setTicketTypes(types);
        if (!ticketTypeId && types.length === 1) setTicketTypeId(types[0].id);
      })
      .catch((e) => setLoadError((e as Error).message));
  }, [open, email.subject, ticketTypeId]);

  const submit = async () => {
    setBusy(true);
    try {
      const html = await getBody("html");
      let attachments: HaloAttachmentInline[] = [];
      let attachWarning: string | undefined;
      if (includeAttachments && attachmentCount > 0) {
        const fetched = await fetchAllAttachments();
        attachments = fetched.attachments.map(toHaloAttachment);
        if (fetched.errors.length > 0) {
          attachWarning = `Some attachments couldn't be included: ${fetched.errors.join("; ")}`;
        }
      }

      const ticket = await createTicket({
        summary,
        details: html,
        client_id: client?.id,
        user_id: contact?.id,
        tickettype_id: ticketTypeId,
        attachments: attachments.length ? attachments : undefined,
        customfields: [
          { name: CUSTOM_FIELD_CONVERSATION_ID, value: email.conversationId },
          { name: CUSTOM_FIELD_MESSAGE_ID, value: email.internetMessageId },
        ],
      });

      // Remember selected ticket type as the new default
      if (ticketTypeId) {
        await setDefaults({ ...getDefaults(), defaultTicketTypeId: ticketTypeId });
      }

      if (attachWarning) {
        onResult("warning", `Created #${ticket.id}, but: ${attachWarning}`);
      } else {
        onResult("success", `Created #${ticket.id}`);
      }
      setOpen(false);
    } catch (e) {
      onResult("error", (e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, d) => setOpen(d.open)}>
      <DialogTrigger disableButtonEnhancement>
        <Button appearance={appearance} icon={<Add24Regular />} className={triggerClass}>
          Create
        </Button>
      </DialogTrigger>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Create ticket from email</DialogTitle>
          <DialogContent>
            {dedupWarning && (
              <MessageBar intent="warning" style={{ marginBottom: 12 }}>
                <MessageBarBody>
                  This conversation is already linked to a ticket. Consider appending instead
                  to keep the thread together.
                </MessageBarBody>
              </MessageBar>
            )}
            <Field label="Summary" required>
              <Input value={summary} onChange={(_, d) => setSummary(d.value)} />
            </Field>

            <Field label="Ticket type" hint="Saved as your default for next time">
              <Combobox
                placeholder={loadError ? "Failed to load types" : "Select…"}
                value={
                  ticketTypeId
                    ? ticketTypes.find((t) => t.id === ticketTypeId)?.name ?? ""
                    : ""
                }
                onOptionSelect={(_, d) =>
                  setTicketTypeId(d.optionValue ? Number(d.optionValue) : undefined)
                }
              >
                {ticketTypes.map((t) => (
                  <Option key={t.id} value={String(t.id)} text={t.name}>
                    {t.name}
                  </Option>
                ))}
              </Combobox>
            </Field>

            <div style={{ marginTop: 12 }}>
              <Switch
                checked={includeAttachments}
                onChange={(_, d) => setIncludeAttachments(d.checked)}
                disabled={attachmentCount === 0}
                label={
                  attachmentCount === 0
                    ? "No attachments"
                    : `Include ${attachmentCount} attachment${attachmentCount === 1 ? "" : "s"}`
                }
              />
            </div>

            <Text size={200} style={{ marginTop: 12, color: tokens.colorNeutralForeground3 }}>
              Client: {client?.name ?? "—"} · Contact: {contact?.name ?? "—"}
              <br />
              Conversation ID stored on the ticket for thread linking on future replies.
            </Text>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => setOpen(false)} disabled={busy}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              onClick={submit}
              disabled={!summary.trim() || busy}
              icon={busy ? <Spinner size="tiny" /> : undefined}
            >
              {busy ? "Creating…" : "Create"}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

// ---------- helpers ----------

function toHaloAttachment(f: FetchedAttachment): HaloAttachmentInline {
  return {
    filename: f.filename,
    data_base64: f.base64,
    contenttype: f.contentType,
    isimage: f.contentType.startsWith("image/"),
  };
}
