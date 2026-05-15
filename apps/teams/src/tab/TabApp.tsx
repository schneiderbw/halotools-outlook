import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Button,
  Field,
  Input,
  Text,
  Spinner,
  Divider,
  Badge,
  Avatar,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { Search24Regular, ArrowClockwise24Regular } from "@fluentui/react-icons";
import { getConfig, setConfig, clearConfig } from "../lib/config";
import { isAuthenticated, signIn, signOut } from "../lib/auth";
import {
  searchUsers,
  findUserByEmail,
  listOpenTicketsForClient,
  listFeed,
} from "../lib/halo-api";
import type {
  HaloUser,
  HaloClient,
  HaloTicket,
  HaloFeedItem,
} from "../lib/types";

type Phase = "needs-config" | "needs-auth" | "ready";

const useStyles = makeStyles({
  root: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
    fontFamily: tokens.fontFamilyBase,
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "12px 20px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  body: {
    flex: 1,
    overflowY: "auto",
    padding: "20px",
    display: "grid",
    // Teams tabs render full-width — a two-column dossier layout works much
    // better than the cramped task pane single column. Falls back to a single
    // column under 720px (Teams mobile / split view).
    gridTemplateColumns: "minmax(280px, 360px) 1fr",
    gap: "20px",
    "@media (max-width: 720px)": {
      gridTemplateColumns: "1fr",
    },
  },
  centeredPane: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "16px",
    padding: "32px",
    maxWidth: "480px",
    margin: "0 auto",
  },
  brand: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
  },
  card: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "16px",
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  cardTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  searchRow: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-end",
  },
  ticketRow: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    padding: "10px 12px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: "pointer",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  ticketTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  ticketMeta: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  feedRow: {
    display: "flex",
    gap: "10px",
    paddingBlock: "10px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
  },
  feedBody: {
    flex: 1,
    minWidth: 0,
  },
  feedNote: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    overflow: "hidden",
    display: "-webkit-box",
    WebkitLineClamp: 3,
    WebkitBoxOrient: "vertical",
  },
  empty: {
    fontStyle: "italic",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
});

export function TabApp() {
  const styles = useStyles();
  const [phase, setPhase] = useState<Phase>(() => initialPhase());

  const refreshPhase = useCallback(() => setPhase(initialPhase()), []);

  if (phase === "needs-config") {
    return (
      <div className={styles.root}>
        <div className={styles.centeredPane}>
          <ConfigForm onConfigured={refreshPhase} />
        </div>
      </div>
    );
  }

  if (phase === "needs-auth") {
    return (
      <div className={styles.root}>
        <div className={styles.centeredPane}>
          <AuthForm onAuthenticated={refreshPhase} onReconfigure={refreshPhase} />
        </div>
      </div>
    );
  }

  return <ReadyTab styles={styles} onSignedOut={refreshPhase} />;
}

function initialPhase(): Phase {
  if (!getConfig()) return "needs-config";
  if (!isAuthenticated()) return "needs-auth";
  return "ready";
}

// ---------- Config form ----------

function ConfigForm({ onConfigured }: { onConfigured: () => void }) {
  const [haloUrl, setHaloUrl] = useState("");
  const [clientId, setClientId] = useState("");
  const [error, setError] = useState<string | undefined>();
  const [saving, setSaving] = useState(false);

  const save = async () => {
    setError(undefined);
    setSaving(true);
    try {
      await setConfig({ haloBaseUrl: haloUrl.trim(), clientId: clientId.trim() });
      onConfigured();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSaving(false);
    }
  };

  return (
    <>
      <Text size={500} weight="semibold">
        Connect HaloPSA
      </Text>
      <Text size={200}>
        Your HaloPSA administrator should register a Halo Connect Application with
        Authorization Code grant, PKCE enabled, redirect URI{" "}
        <strong>https://tools.iusehalo.com/teams/auth/callback.html</strong>, and{" "}
        <strong>https://tools.iusehalo.com</strong> in the CORS whitelist.
      </Text>
      <Field label="HaloPSA URL" required hint="e.g. https://halo.yourcompany.com">
        <Input
          value={haloUrl}
          onChange={(_, d) => setHaloUrl(d.value)}
          placeholder="https://halo.yourcompany.com"
          autoComplete="off"
        />
      </Field>
      <Field label="Client ID" required hint="From the Halo Connect Application">
        <Input
          value={clientId}
          onChange={(_, d) => setClientId(d.value)}
          placeholder="00000000-0000-0000-0000-000000000000"
          autoComplete="off"
        />
      </Field>
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      <Button
        appearance="primary"
        disabled={!haloUrl.trim() || !clientId.trim() || saving}
        onClick={save}
      >
        {saving ? "Saving…" : "Continue"}
      </Button>
    </>
  );
}

// ---------- Auth form ----------

function AuthForm({
  onAuthenticated,
  onReconfigure,
}: {
  onAuthenticated: () => void;
  onReconfigure: () => void;
}) {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | undefined>();

  const start = async () => {
    setBusy(true);
    setError(undefined);
    try {
      await signIn();
      onAuthenticated();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const wipe = async () => {
    await clearConfig();
    onReconfigure();
  };

  return (
    <>
      <Text size={500} weight="semibold">
        Sign in to HaloPSA
      </Text>
      <Text size={200}>
        Connect this Teams app to your HaloPSA tenant. A popup will open Halo's sign-in
        page; once you authorize the app, the popup closes and you're returned here.
      </Text>
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      <div style={{ display: "flex", gap: 8 }}>
        <Button appearance="primary" disabled={busy} onClick={start}>
          {busy ? "Signing in…" : "Connect to Halo"}
        </Button>
        <Button appearance="subtle" disabled={busy} onClick={wipe}>
          Change tenant
        </Button>
      </div>
    </>
  );
}

// ---------- Ready / dashboard ----------

interface ReadyProps {
  styles: ReturnType<typeof useStyles>;
  onSignedOut: () => void;
}

function ReadyTab({ styles, onSignedOut }: ReadyProps) {
  const cfg = useMemo(() => getConfig(), []);
  const [query, setQuery] = useState("");
  const [searching, setSearching] = useState(false);
  const [results, setResults] = useState<HaloUser[]>([]);
  const [selected, setSelected] = useState<HaloUser | undefined>();
  const [client, setClient] = useState<HaloClient | undefined>();
  const [openTickets, setOpenTickets] = useState<HaloTicket[]>([]);
  const [feed, setFeed] = useState<HaloFeedItem[]>([]);
  const [loadingDetail, setLoadingDetail] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [refreshTick, setRefreshTick] = useState(0);

  const runSearch = useCallback(async () => {
    setError(undefined);
    const q = query.trim();
    if (!q) {
      setResults([]);
      return;
    }
    setSearching(true);
    try {
      // If the query looks like an email, prefer exact email lookup; otherwise
      // fall back to broad user search. Mirrors how the Outlook side reconciles
      // an incoming email's sender to a contact.
      if (q.includes("@")) {
        const exact = await findUserByEmail(q);
        if (exact) {
          setResults([exact]);
        } else {
          setResults(await searchUsers(q));
        }
      } else {
        setResults(await searchUsers(q));
      }
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSearching(false);
    }
  }, [query]);

  // Whenever a contact is selected, fetch their client's open tickets and the
  // activity feed for that client. Both are best-effort — feed is decorative.
  useEffect(() => {
    if (!selected) {
      setClient(undefined);
      setOpenTickets([]);
      setFeed([]);
      return;
    }
    if (!selected.client_id) {
      setClient(undefined);
      setOpenTickets([]);
      setFeed([]);
      return;
    }
    let cancelled = false;
    setLoadingDetail(true);
    const clientStub: HaloClient = {
      id: selected.client_id,
      name: selected.client_name ?? "",
    };
    setClient(clientStub);

    (async () => {
      try {
        const [tickets, feedItems] = await Promise.all([
          listOpenTicketsForClient(selected.client_id!).catch(() => [] as HaloTicket[]),
          listFeed({ user_id: selected.id }, 25).catch(() => [] as HaloFeedItem[]),
        ]);
        if (cancelled) return;
        setOpenTickets(tickets);
        setFeed(feedItems);
      } catch (e) {
        if (!cancelled) setError((e as Error).message);
      } finally {
        if (!cancelled) setLoadingDetail(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [selected?.id, refreshTick]);

  const handleSignOut = useCallback(async () => {
    await signOut();
    onSignedOut();
  }, [onSignedOut]);

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text className={styles.brand}>HaloPSA</Text>
        <div style={{ display: "flex", gap: 8 }}>
          <Button
            appearance="subtle"
            icon={<ArrowClockwise24Regular />}
            onClick={() => setRefreshTick((n) => n + 1)}
            aria-label="Refresh"
          />
          <Button appearance="subtle" onClick={handleSignOut}>
            Sign out
          </Button>
        </div>
      </div>

      <div className={styles.body}>
        <div className={styles.card}>
          <Text className={styles.cardTitle}>Find a contact</Text>
          <div className={styles.searchRow}>
            <Field style={{ flex: 1 }}>
              <Input
                value={query}
                onChange={(_, d) => setQuery(d.value)}
                placeholder="Name or email"
                contentBefore={<Search24Regular />}
                onKeyDown={(e) => {
                  if (e.key === "Enter") void runSearch();
                }}
              />
            </Field>
            <Button appearance="primary" onClick={() => void runSearch()} disabled={searching}>
              {searching ? "…" : "Search"}
            </Button>
          </div>
          {error && (
            <MessageBar intent="error">
              <MessageBarBody>{error}</MessageBarBody>
            </MessageBar>
          )}
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            {results.length === 0 && !searching && (
              <Text className={styles.empty}>
                Search by name or email to load a contact's open tickets and feed.
              </Text>
            )}
            {results.map((u) => (
              <div
                key={u.id}
                className={styles.ticketRow}
                onClick={() => setSelected(u)}
                style={
                  selected?.id === u.id
                    ? { borderColor: tokens.colorBrandStroke1, borderWidth: 2 }
                    : undefined
                }
              >
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <Avatar name={u.name} size={24} />
                  <Text weight="semibold">{u.name}</Text>
                </div>
                <Text className={styles.ticketMeta}>
                  {u.emailaddress ?? "—"}
                  {u.client_name ? ` · ${u.client_name}` : ""}
                </Text>
              </div>
            ))}
          </div>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <div className={styles.card}>
            <Text className={styles.cardTitle}>
              {selected
                ? `${selected.name}'s open tickets${client?.name ? ` · ${client.name}` : ""}`
                : "Open tickets"}
            </Text>
            {!selected && (
              <Text className={styles.empty}>Select a contact to load their tickets.</Text>
            )}
            {loadingDetail && (
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <Spinner size="extra-small" /> <Text size={200}>Loading…</Text>
              </div>
            )}
            {!loadingDetail && selected && openTickets.length === 0 && (
              <Text className={styles.empty}>No open tickets for this contact's client.</Text>
            )}
            {openTickets.map((t) => (
              <a
                key={t.id}
                href={cfg ? `${cfg.haloBaseUrl}/ticket?id=${t.id}` : "#"}
                target="_blank"
                rel="noreferrer"
                className={styles.ticketRow}
                style={{ textDecoration: "none", color: "inherit" }}
              >
                <Text className={styles.ticketTitle}>
                  #{t.id} · {t.summary}
                </Text>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {t.statusname && <Badge appearance="tint">{t.statusname}</Badge>}
                  {t.priorityname && <Badge appearance="outline">{t.priorityname}</Badge>}
                  {t.agent_name && <Badge appearance="outline">{t.agent_name}</Badge>}
                </div>
              </a>
            ))}
          </div>

          <div className={styles.card}>
            <Text className={styles.cardTitle}>Recent activity</Text>
            {!selected && (
              <Text className={styles.empty}>Select a contact to load their activity feed.</Text>
            )}
            {selected && feed.length === 0 && !loadingDetail && (
              <Text className={styles.empty}>No recent activity.</Text>
            )}
            {feed.map((f) => (
              <div key={f.id} className={styles.feedRow}>
                <Avatar name={f.who_name ?? "?"} size={28} />
                <div className={styles.feedBody}>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      gap: 8,
                    }}
                  >
                    <Text size={200} weight="semibold">
                      {f.who_name ?? "Unknown"}
                    </Text>
                    <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>
                      {formatWhen(f.datetime)}
                    </Text>
                  </div>
                  {f.outcome && (
                    <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>
                      {f.outcome}
                    </Text>
                  )}
                  {f.note && <div className={styles.feedNote}>{stripHtml(f.note)}</div>}
                </div>
              </div>
            ))}
          </div>

          <Divider />
          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>
            Connected to {cfg?.haloBaseUrl}
          </Text>
        </div>
      </div>
    </div>
  );
}

function formatWhen(iso: string | undefined): string {
  if (!iso) return "";
  try {
    const d = new Date(iso);
    return d.toLocaleString();
  } catch {
    return iso;
  }
}

function stripHtml(input: string): string {
  // Halo notes are often HTML; the feed card only has ~3 lines, so strip tags
  // and collapse whitespace for a clean preview rather than dangerously rendering.
  return input.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}
