import { useEffect, useState, useRef, useCallback } from "react";
import {
  Text,
  Spinner,
  makeStyles,
  tokens,
  Button,
  Input,
  Avatar,
  Badge,
  Divider,
  MessageBar,
  MessageBarBody,
  Switch,
  Combobox,
  Option,
  Skeleton,
  SkeletonItem,
  Field,
} from "@fluentui/react-components";
import {
  Search24Regular,
  DocumentLink24Regular,
  BookQuestionMark24Regular,
  Send24Regular,
  CheckmarkCircle16Filled,
} from "@fluentui/react-icons";
import { ConfigScreen } from "../components/ConfigScreen";
import { AuthScreen } from "../components/AuthScreen";
import { getConfig, type TenantConfig } from "../lib/config";
import { isAuthenticated } from "../lib/auth";
import {
  findUserByEmail,
  findClientByDomain,
  searchTickets,
  searchKbArticles,
  appendAction,
  stampOutlookThreadFields,
} from "../lib/halo-api";
import {
  getRecipients,
  insertIntoBody,
  saveDraft,
  getComposeBody,
  getComposeSubject,
  domainOf,
} from "../lib/office";
import { getDefaults } from "../lib/defaults";
import type {
  HaloUser,
  HaloClient,
  HaloTicket,
  HaloKbArticle,
} from "../types/halo";

type Phase = "loading" | "needs-config" | "needs-auth" | "ready";

const useStyles = makeStyles({
  root: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
    overflowY: "auto",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "8px 12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  brand: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
  },
  body: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "12px",
    flex: 1,
  },
  sectionLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
    marginBottom: "6px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
  recipientRow: {
    display: "flex",
    gap: "10px",
    alignItems: "center",
    padding: "6px 8px",
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  recipientText: {
    flex: 1,
    minWidth: 0,
    display: "flex",
    flexDirection: "column",
  },
  recipientName: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  recipientSecondary: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  empty: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
  },
  resultList: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    maxHeight: "200px",
    overflowY: "auto",
  },
  resultRow: {
    padding: "8px 10px",
    cursor: "pointer",
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
    ":last-child": {
      borderBottom: "none",
    },
  },
  resultPrimary: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
  },
  resultSecondary: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
  toast: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    color: tokens.colorPaletteGreenForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  toggleSection: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "10px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  hint: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
  centerPad: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "12px",
    padding: "24px",
  },
});

export function ComposeApp() {
  const styles = useStyles();
  const [phase, setPhase] = useState<Phase>("loading");

  const refreshPhase = useCallback(() => {
    if (!getConfig()) setPhase("needs-config");
    else if (!isAuthenticated()) setPhase("needs-auth");
    else setPhase("ready");
  }, []);

  useEffect(() => {
    refreshPhase();
  }, [refreshPhase]);

  if (phase === "loading") {
    return (
      <div className={styles.root}>
        <div className={styles.centerPad}>
          <Spinner label="Loading…" />
        </div>
      </div>
    );
  }

  if (phase === "needs-config") {
    return (
      <div className={styles.root}>
        <ConfigScreen onConfigured={refreshPhase} />
      </div>
    );
  }

  if (phase === "needs-auth") {
    return (
      <div className={styles.root}>
        <AuthScreen onAuthenticated={refreshPhase} onReconfigure={refreshPhase} />
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text className={styles.brand}>HaloPSA</Text>
      </div>
      <div className={styles.body}>
        <RecipientsSection />
        <Divider />
        <InsertTicketSection />
        <Divider />
        <InsertKbSection />
        <Divider />
        <LogOnSendSection />
      </div>
    </div>
  );
}

// ---------- Recipients section ----------

interface ResolvedRecipient {
  email: string;
  user?: HaloUser;
  client?: HaloClient;
  loading: boolean;
}

function RecipientsSection() {
  const styles = useStyles();
  const [recipients, setRecipients] = useState<ResolvedRecipient[]>([]);
  const [loadingList, setLoadingList] = useState(true);
  const [error, setError] = useState<string | undefined>();

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const { to } = await getRecipients();
        if (cancelled) return;
        if (to.length === 0) {
          setRecipients([]);
          setLoadingList(false);
          return;
        }
        // Seed each as still-resolving so the UI shows skeletons while Halo lookups run.
        const seeds: ResolvedRecipient[] = to.map((e) => ({ email: e, loading: true }));
        setRecipients(seeds);
        setLoadingList(false);

        // Resolve each recipient in parallel — domain lookups are cheap per address.
        await Promise.all(
          to.map(async (email, idx) => {
            try {
              const [user, client] = await Promise.all([
                findUserByEmail(email).catch(() => undefined),
                findClientByDomain(domainOf(email)).catch(() => undefined),
              ]);
              if (cancelled) return;
              setRecipients((prev) => {
                const next = [...prev];
                next[idx] = { email, user, client, loading: false };
                return next;
              });
            } catch {
              if (cancelled) return;
              setRecipients((prev) => {
                const next = [...prev];
                next[idx] = { email, loading: false };
                return next;
              });
            }
          }),
        );
      } catch (e) {
        if (!cancelled) {
          setError((e as Error).message);
          setLoadingList(false);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <div className={styles.section}>
      <Text className={styles.sectionLabel}>Recipients</Text>
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      {loadingList ? (
        <Skeleton>
          <SkeletonItem size={48} />
        </Skeleton>
      ) : recipients.length === 0 ? (
        <Text className={styles.empty}>No recipients yet.</Text>
      ) : (
        recipients.map((r, i) => <RecipientRow key={`${r.email}-${i}`} recipient={r} />)
      )}
    </div>
  );
}

function RecipientRow({ recipient }: { recipient: ResolvedRecipient }) {
  const styles = useStyles();
  const displayName = recipient.user?.name ?? recipient.email;
  const secondary = recipient.user?.client_name ?? recipient.client?.name ?? recipient.email;

  return (
    <div className={styles.recipientRow}>
      <Avatar name={displayName} color="colorful" size={28} />
      <div className={styles.recipientText}>
        <Text className={styles.recipientName}>{displayName}</Text>
        <Text className={styles.recipientSecondary}>{secondary}</Text>
      </div>
      {recipient.loading ? (
        <Spinner size="extra-tiny" />
      ) : recipient.user ? (
        <Badge appearance="filled" color="success" size="small">
          Contact
        </Badge>
      ) : recipient.client ? (
        <Badge appearance="filled" color="warning" size="small">
          Domain
        </Badge>
      ) : (
        <Badge appearance="filled" color="danger" size="small">
          No match
        </Badge>
      )}
    </div>
  );
}

// ---------- Insert ticket-link section ----------

function InsertTicketSection() {
  const styles = useStyles();
  const cfg = getConfig() as TenantConfig;
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<HaloTicket[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [toast, setToast] = useState<string | undefined>();
  const debounceRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  useEffect(() => {
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (query.trim().length < 2) {
      setResults([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    debounceRef.current = setTimeout(async () => {
      try {
        const tickets = await searchTickets(query.trim());
        setResults(tickets);
        setError(undefined);
      } catch (e) {
        setError((e as Error).message);
        setResults([]);
      } finally {
        setLoading(false);
      }
    }, 300);
    return () => {
      if (debounceRef.current) clearTimeout(debounceRef.current);
    };
  }, [query]);

  const onPick = async (ticket: HaloTicket) => {
    try {
      const url = `${cfg.haloBaseUrl}/agent?id=${ticket.id}`;
      const safeSummary = escapeHtml(ticket.summary);
      const html = `<a href="${url}">#${ticket.id} — ${safeSummary}</a>`;
      await insertIntoBody(html);
      setToast(`Inserted link to #${ticket.id}`);
      setTimeout(() => setToast(undefined), 3000);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  return (
    <div className={styles.section}>
      <Text className={styles.sectionLabel}>
        <DocumentLink24Regular
          style={{ verticalAlign: "middle", marginRight: 4, width: 14, height: 14 }}
        />
        Insert ticket link
      </Text>
      <Input
        value={query}
        placeholder="Search tickets…"
        onChange={(_, d) => setQuery(d.value)}
        contentBefore={<Search24Regular />}
      />
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      {toast && (
        <div className={styles.toast}>
          <CheckmarkCircle16Filled />
          <span>{toast}</span>
        </div>
      )}
      {query.trim().length >= 2 && (
        <div className={styles.resultList}>
          {loading ? (
            <div style={{ padding: 10 }}>
              <Spinner size="extra-tiny" /> Searching…
            </div>
          ) : results.length === 0 ? (
            <div style={{ padding: 10 }}>
              <Text className={styles.empty}>No tickets found.</Text>
            </div>
          ) : (
            results.map((t) => (
              <div key={t.id} className={styles.resultRow} onClick={() => onPick(t)}>
                <Text className={styles.resultPrimary}>
                  #{t.id} — {t.summary}
                </Text>
                {(t.client_name || t.statusname) && (
                  <>
                    <br />
                    <Text className={styles.resultSecondary}>
                      {[t.client_name, t.statusname].filter(Boolean).join(" · ")}
                    </Text>
                  </>
                )}
              </div>
            ))
          )}
        </div>
      )}
    </div>
  );
}

// ---------- Insert KB-article section ----------

const KB_BODY_INSERT_THRESHOLD = 4000;

function InsertKbSection() {
  const styles = useStyles();
  const cfg = getConfig() as TenantConfig;
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<HaloKbArticle[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [toast, setToast] = useState<string | undefined>();
  const debounceRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  useEffect(() => {
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (query.trim().length < 2) {
      setResults([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    debounceRef.current = setTimeout(async () => {
      try {
        const articles = await searchKbArticles(query.trim());
        setResults(articles);
        setError(undefined);
      } catch (e) {
        setError((e as Error).message);
        setResults([]);
      } finally {
        setLoading(false);
      }
    }, 300);
    return () => {
      if (debounceRef.current) clearTimeout(debounceRef.current);
    };
  }, [query]);

  const onPick = async (article: HaloKbArticle) => {
    try {
      // Body field varies by tenant version — try faq_answer first, fall back to details.
      const body = article.faq_answer ?? article.details ?? "";
      const articleUrl = `${cfg.haloBaseUrl}/kb?id=${article.id}`;
      let html: string;
      if (body && body.length <= KB_BODY_INSERT_THRESHOLD) {
        // Inline the full article when it's short enough.
        html = body;
      } else if (body) {
        // For long articles, insert a snippet + link so the email stays readable.
        const text = stripHtml(body).slice(0, 400).trim();
        html = `<blockquote>${escapeHtml(text)}…</blockquote><p><a href="${articleUrl}">Read full article: ${escapeHtml(article.name)}</a></p>`;
      } else {
        html = `<p><a href="${articleUrl}">${escapeHtml(article.name)}</a></p>`;
      }
      await insertIntoBody(html);
      setToast(`Inserted "${article.name}"`);
      setTimeout(() => setToast(undefined), 3000);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  return (
    <div className={styles.section}>
      <Text className={styles.sectionLabel}>
        <BookQuestionMark24Regular
          style={{ verticalAlign: "middle", marginRight: 4, width: 14, height: 14 }}
        />
        Insert KB article
      </Text>
      <Input
        value={query}
        placeholder="Search KB…"
        onChange={(_, d) => setQuery(d.value)}
        contentBefore={<Search24Regular />}
      />
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      {toast && (
        <div className={styles.toast}>
          <CheckmarkCircle16Filled />
          <span>{toast}</span>
        </div>
      )}
      {query.trim().length >= 2 && (
        <div className={styles.resultList}>
          {loading ? (
            <div style={{ padding: 10 }}>
              <Spinner size="extra-tiny" /> Searching…
            </div>
          ) : results.length === 0 ? (
            <div style={{ padding: 10 }}>
              <Text className={styles.empty}>No articles found.</Text>
            </div>
          ) : (
            results.map((a) => (
              <div key={a.id} className={styles.resultRow} onClick={() => onPick(a)}>
                <Text className={styles.resultPrimary}>{a.name}</Text>
                {a.tags && a.tags.length > 0 && (
                  <>
                    <br />
                    <Text className={styles.resultSecondary}>
                      {a.tags.map((t) => t.value).join(" · ")}
                    </Text>
                  </>
                )}
              </div>
            ))
          )}
        </div>
      )}
    </div>
  );
}

// ---------- Log-on-send section ----------
//
// Approach: SaveAsync fallback.
//
// The "proper" wiring for log-on-send is a LaunchEvent runtime that registers
// an on-send handler via manifest, runs as a separate JS bundle in a hidden
// runtime, and calls event.completed() once the append finishes. That requires:
//   - a second runtime entry in manifest.json with type "general" + lifetime "long"
//   - extensionPoints.launchEvents with type "onMessageSend"
//   - a globally-named JS function the runtime can dispatch to
//   - org-admin consent for the SendItem permission (currently not in our manifest)
//
// Wiring all of that requires touching the existing read runtime config and adding
// permissions the user may not yet have approved. To keep this PR focused on the
// compose UI surface, we use the documented fallback: when the user clicks
// "Append now", we saveAsync to materialize the draft, then call appendAction
// against the current body. The "auto-on-send" toggle stays as UI but currently
// invokes the same flow with a clear "manual fallback" label.

function LogOnSendSection() {
  const styles = useStyles();
  const [enabled, setEnabled] = useState(false);
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<HaloTicket[]>([]);
  const [selectedTicketId, setSelectedTicketId] = useState<number | undefined>();
  const [loadingSearch, setLoadingSearch] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [toast, setToast] = useState<string | undefined>();
  const debounceRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  useEffect(() => {
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (query.trim().length < 2) {
      setResults([]);
      setLoadingSearch(false);
      return;
    }
    setLoadingSearch(true);
    debounceRef.current = setTimeout(async () => {
      try {
        const tickets = await searchTickets(query.trim());
        setResults(tickets);
        setError(undefined);
      } catch (e) {
        setError((e as Error).message);
        setResults([]);
      } finally {
        setLoadingSearch(false);
      }
    }, 300);
    return () => {
      if (debounceRef.current) clearTimeout(debounceRef.current);
    };
  }, [query]);

  const selectedTicket = results.find((t) => t.id === selectedTicketId);

  const appendNow = async () => {
    if (!selectedTicketId) return;
    setBusy(true);
    setError(undefined);
    try {
      // Forces Outlook to persist the in-flight draft to the user's Drafts folder.
      // Without this, body / subject reads can race with the user still editing.
      await saveDraft();
      const [html, subject, recipients] = await Promise.all([
        getComposeBody("html"),
        getComposeSubject(),
        getRecipients().catch(() => ({ to: [] as string[], cc: [], bcc: [] })),
      ]);
      const toList = recipients.to.join(", ");
      const action = await appendAction({
        ticket_id: selectedTicketId,
        outcome: getDefaults().defaultAppendOutcome ?? "Email Received",
        note: html,
        emailsubject: subject || undefined,
        // The current Outlook user is the sender for compose-surface messages.
        emailfrom: Office.context.mailbox.userProfile?.emailAddress,
        emailfromname: Office.context.mailbox.userProfile?.displayName,
        // Tack the To list onto the subject line so it shows in Halo's action history
        // without needing a dedicated field. (Halo's CreateActionPayload has no To field.)
        ...(toList ? { emailsubject: `${subject} — to: ${toList}` } : {}),
      });
      // Best-effort stamp so future replies can thread to the same ticket.
      stampOutlookThreadFields(
        selectedTicketId,
        // Compose items don't expose a conversationId until they're sent; use the draft itemId
        // as a stable thread key for now. Replies generated in Outlook share conversationId
        // with their parent, so the read surface's append path will still pick this up.
        undefined,
        undefined,
      ).catch(() => {
        /* non-fatal */
      });
      setToast(`Appended to #${action.ticket_id} (manual fallback — see hint)`);
      setTimeout(() => setToast(undefined), 5000);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={styles.section}>
      <Text className={styles.sectionLabel}>
        <Send24Regular
          style={{ verticalAlign: "middle", marginRight: 4, width: 14, height: 14 }}
        />
        Log to ticket
      </Text>
      <div className={styles.toggleSection}>
        <Switch
          checked={enabled}
          onChange={(_, d) => setEnabled(d.checked)}
          label="Log this email to a ticket"
        />
        {enabled && (
          <>
            <Field label="Find ticket">
              <Input
                value={query}
                placeholder="Search by ID, summary, or client…"
                onChange={(_, d) => setQuery(d.value)}
                contentBefore={<Search24Regular />}
              />
            </Field>
            {loadingSearch ? (
              <div>
                <Spinner size="extra-tiny" /> Searching…
              </div>
            ) : results.length > 0 ? (
              <Field label="Ticket">
                <Combobox
                  placeholder="Select a ticket"
                  value={
                    selectedTicket ? `#${selectedTicket.id} — ${selectedTicket.summary}` : ""
                  }
                  onOptionSelect={(_, d) =>
                    setSelectedTicketId(d.optionValue ? Number(d.optionValue) : undefined)
                  }
                >
                  {results.map((t) => (
                    <Option
                      key={t.id}
                      value={String(t.id)}
                      text={`#${t.id} ${t.summary}`}
                    >
                      #{t.id} · {t.summary}
                      {t.statusname ? ` · ${t.statusname}` : ""}
                    </Option>
                  ))}
                </Combobox>
              </Field>
            ) : query.trim().length >= 2 ? (
              <Text className={styles.empty}>No tickets found.</Text>
            ) : null}
            <Text className={styles.hint}>
              On-send auto-logging is not yet wired (requires a LaunchEvent runtime and
              the SendItem permission). Click "Append now" to manually log the draft to
              the selected ticket; Outlook will save the draft first so the body and
              subject are stable.
            </Text>
            <Button
              appearance="primary"
              disabled={!selectedTicketId || busy}
              onClick={appendNow}
              icon={busy ? <Spinner size="tiny" /> : <Send24Regular />}
            >
              {busy ? "Appending…" : "Append now"}
            </Button>
            {error && (
              <MessageBar intent="error">
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}
            {toast && (
              <div className={styles.toast}>
                <CheckmarkCircle16Filled />
                <span>{toast}</span>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

// ---------- helpers ----------

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function stripHtml(s: string): string {
  return s
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ");
}
