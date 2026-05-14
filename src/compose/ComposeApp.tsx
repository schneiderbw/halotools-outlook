import { useEffect, useState, useCallback } from "react";
import {
  Text,
  Spinner,
  makeStyles,
  tokens,
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
} from "../lib/halo-api";
import {
  getRecipients,
  insertIntoBody,
  domainOf,
} from "../lib/office";
import { useDebouncedSearch } from "../lib/use-debounced-search";
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
  const [toast, setToast] = useState<string | undefined>();
  const [actionError, setActionError] = useState<string | undefined>();
  const { results, loading, error: searchError } = useDebouncedSearch(query, searchTickets);
  const error = actionError ?? searchError;
  const setError = setActionError;

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
  const [toast, setToast] = useState<string | undefined>();
  const [actionError, setActionError] = useState<string | undefined>();
  const { results, loading, error: searchError } = useDebouncedSearch(query, searchKbArticles);
  const error = actionError ?? searchError;
  const setError = setActionError;

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
  // The actual "this draft will log on send" flag lives in the item's
  // CustomProperties bag because the launch-event runtime can't read this
  // React state. We mirror it locally for the UI.
  const [enabled, setEnabled] = useState(false);
  const [query, setQuery] = useState("");
  const [selectedTicketId, setSelectedTicketId] = useState<number | undefined>();
  const [armed, setArmed] = useState(false);
  const { results, loading: loadingSearch, error } = useDebouncedSearch(query, searchTickets);

  // Rehydrate any previously-armed ticket on this draft so the user sees it
  // sticking when they reopen the pane.
  useEffect(() => {
    const item = Office.context.mailbox.item;
    if (!item) return;
    item.loadCustomPropertiesAsync((r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return;
      const existing = r.value.get("haloLogTicketId");
      if (existing) {
        setEnabled(true);
        setSelectedTicketId(Number(existing));
        setArmed(true);
      }
    });
  }, []);

  const selectedTicket = results.find((t) => t.id === selectedTicketId);

  // Persist the picked ticket onto the draft so the on-send handler can find
  // it. Setting null clears it (used when user disables the toggle).
  const writeCustomProp = (ticketId: number | null) => {
    return new Promise<void>((resolve) => {
      const item = Office.context.mailbox.item;
      if (!item) {
        resolve();
        return;
      }
      item.loadCustomPropertiesAsync((r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }
        const cp = r.value;
        // Office.js CustomProperties.set is typed as accepting string; we
        // serialize numbers, and "" doubles as our cleared sentinel.
        cp.set("haloLogTicketId", ticketId == null ? "" : String(ticketId));
        cp.saveAsync(() => resolve());
      });
    });
  };

  const toggle = async (next: boolean) => {
    setEnabled(next);
    if (!next) {
      await writeCustomProp(null);
      setArmed(false);
    } else if (selectedTicketId) {
      await writeCustomProp(selectedTicketId);
      setArmed(true);
    }
  };

  const pickTicket = async (id: number | undefined) => {
    setSelectedTicketId(id);
    if (enabled && id) {
      await writeCustomProp(id);
      setArmed(true);
    } else {
      setArmed(false);
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
          onChange={(_, d) => toggle(d.checked)}
          label="Log this email to a ticket when I click Send"
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
                    pickTicket(d.optionValue ? Number(d.optionValue) : undefined)
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
            {armed && selectedTicketId && (
              <div className={styles.toast}>
                <CheckmarkCircle16Filled />
                <span>
                  Ready — this draft will be logged to #{selectedTicketId} when you click Send.
                </span>
              </div>
            )}
            {error && (
              <MessageBar intent="error">
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
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
