import { useEffect, useState } from "react";
import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Input,
  Textarea,
  Field,
  MessageBar,
  MessageBarBody,
  Spinner,
  Switch,
  tokens,
  Text,
} from "@fluentui/react-components";
import type { HaloUser, HaloClient } from "@iusehalo/halo-api";
import { createCRMNote } from "@iusehalo/halo-api";

interface Props {
  open: boolean;
  contact?: HaloUser;
  client?: HaloClient;
  onClose: () => void;
}

type Scope = "user" | "client";

export function LogNoteDialog({ open, contact, client, onClose }: Props) {
  const [scope, setScope] = useState<Scope>(contact ? "user" : "client");
  const [subject, setSubject] = useState("");
  const [note, setNote] = useState("");
  const [minutes, setMinutes] = useState("");
  const [recordTime, setRecordTime] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [done, setDone] = useState(false);

  useEffect(() => {
    if (!open) return;
    setScope(contact ? "user" : "client");
    setSubject("");
    setNote("");
    setMinutes("");
    setRecordTime(false);
    setBusy(false);
    setError(undefined);
    setDone(false);
  }, [open, contact]);

  const submit = async () => {
    if (!subject.trim() || !note.trim()) {
      setError("Subject and note are required.");
      return;
    }
    setBusy(true);
    setError(undefined);
    try {
      const mins = parseFloat(minutes);
      const hours = recordTime && Number.isFinite(mins) && mins > 0 ? mins / 60 : undefined;
      await createCRMNote({
        client_id: scope === "client" ? client?.id : undefined,
        user_id: scope === "user" ? contact?.id : undefined,
        subject: subject.trim(),
        note: note.trim(),
        timetaken: hours,
        hide_time_taken: !recordTime,
      });
      setDone(true);
      // Auto-close after a beat so the user sees the green confirmation.
      setTimeout(onClose, 900);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const targetLabel =
    scope === "user"
      ? contact
        ? `${contact.name} (contact)`
        : "this contact"
      : client
      ? `${client.name} (client)`
      : "this client";

  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Log a note</DialogTitle>
          <DialogContent>
            {contact && client && (
              <Field label="Attach to">
                <div style={{ display: "flex", gap: 6 }}>
                  <Button
                    size="small"
                    appearance={scope === "user" ? "primary" : "secondary"}
                    onClick={() => setScope("user")}
                  >
                    Contact
                  </Button>
                  <Button
                    size="small"
                    appearance={scope === "client" ? "primary" : "secondary"}
                    onClick={() => setScope("client")}
                  >
                    Client
                  </Button>
                </div>
              </Field>
            )}

            <Text
              style={{
                color: tokens.colorNeutralForeground3,
                fontSize: tokens.fontSizeBase200,
                marginTop: 4,
              }}
            >
              Will appear on the timeline for <strong>{targetLabel}</strong>.
            </Text>

            <Field label="Subject" required style={{ marginTop: 10 }}>
              <Input
                value={subject}
                onChange={(_, d) => setSubject(d.value)}
                placeholder="What happened"
              />
            </Field>

            <Field label="Note" required style={{ marginTop: 10 }}>
              <Textarea
                value={note}
                onChange={(_, d) => setNote(d.value)}
                rows={6}
                placeholder="Details, outcome, follow-ups…"
              />
            </Field>

            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                marginTop: 10,
              }}
            >
              <Switch
                checked={recordTime}
                onChange={(_, d) => setRecordTime(d.checked)}
                label="Record time spent"
              />
              {recordTime && (
                <Input
                  type="number"
                  value={minutes}
                  onChange={(_, d) => setMinutes(d.value)}
                  placeholder="Minutes"
                  style={{ width: 120 }}
                />
              )}
            </div>

            {error && (
              <MessageBar intent="error" style={{ marginTop: 10 }}>
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}
            {done && (
              <MessageBar intent="success" style={{ marginTop: 10 }}>
                <MessageBarBody>Note logged.</MessageBarBody>
              </MessageBar>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose} disabled={busy}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              onClick={submit}
              disabled={busy || done || !subject.trim() || !note.trim()}
              icon={busy ? <Spinner size="tiny" /> : undefined}
            >
              {busy ? "Saving…" : done ? "Saved" : "Log note"}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
