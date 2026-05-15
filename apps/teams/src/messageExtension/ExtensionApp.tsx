import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Button,
  Input,
  Field,
  Text,
  Spinner,
  TabList,
  Tab,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { dialog } from "@microsoft/teams-js";
import { getConfig } from "../lib/config";
import { isAuthenticated, signIn } from "../lib/auth";
import { searchTickets, searchCannedText } from "../lib/halo-api";
import type { HaloTicket, HaloCannedText } from "../lib/types";

// This page renders inside a Teams task module opened from the message-extension
// action. When the user picks a ticket or canned text, we hand the result back
// to Teams via dialog.url.submit(...). Teams forwards that payload to the
// caller (the chat compose box), which uses it as the inserted content.

type Tab = "ticket" | "canned";

const useStyles = makeStyles({
  root: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    padding: "16px",
    gap: "12px",
    fontFamily: tokens.fontFamilyBase,
  },
  searchRow: {
    display: "flex",
    gap: "8px",
  },
  results: {
    flex: 1,
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "8px",
  },
  row: {
    padding: "8px 10px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: "pointer",
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  empty: {
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
    fontSize: tokens.fontSizeBase200,
  },
  footer: {
    display: "flex",
    justifyContent: "space-between",
    gap: "8px",
  },
});

export function ExtensionApp() {
  const styles = useStyles();
  const cfg = useMemo(() => getConfig(), []);
  const [tab, setTab] = useState<Tab>("ticket");
  const [authed, setAuthed] = useState(isAuthenticated());

  if (!cfg) {
    return (
      <div className={styles.root}>
        <Text size={500} weight="semibold">
          HaloPSA not configured
        </Text>
        <Text size={200}>
          Open the Halo tab in Teams first and complete the setup. The message
          extension shares its tenant config with the tab.
        </Text>
      </div>
    );
  }

  if (!authed) {
    return (
      <div className={styles.root}>
        <Text size={500} weight="semibold">
          Sign in to HaloPSA
        </Text>
        <Text size={200}>This message extension needs to authenticate first.</Text>
        <Button
          appearance="primary"
          onClick={async () => {
            try {
              await signIn();
              setAuthed(true);
            } catch (e) {
              alert((e as Error).message);
            }
          }}
        >
          Connect to Halo
        </Button>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as Tab)}>
        <Tab value="ticket">Ticket link</Tab>
        <Tab value="canned">Canned text</Tab>
      </TabList>
      {tab === "ticket" ? (
        <TicketPicker styles={styles} haloBaseUrl={cfg.haloBaseUrl} />
      ) : (
        <CannedPicker styles={styles} />
      )}
    </div>
  );
}

// ---------- Ticket picker ----------

function TicketPicker({
  styles,
  haloBaseUrl,
}: {
  styles: ReturnType<typeof useStyles>;
  haloBaseUrl: string;
}) {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<HaloTicket[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();

  const run = useCallback(async () => {
    const q = query.trim();
    if (!q) return;
    setLoading(true);
    setError(undefined);
    try {
      setResults(await searchTickets(q));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  }, [query]);

  const insert = (t: HaloTicket) => {
    const url = `${haloBaseUrl}/ticket?id=${t.id}`;
    // Build a minimal Adaptive Card v1.4 — Teams accepts these directly from
    // a message-extension action and renders them inline in the chat. Falls
    // back to plain HTML if the host can't render cards.
    const card = {
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: `HaloPSA Ticket #${t.id}`,
          weight: "Bolder",
          size: "Medium",
        },
        {
          type: "TextBlock",
          text: t.summary,
          wrap: true,
        },
        ...(t.statusname || t.client_name
          ? [
              {
                type: "TextBlock",
                text: [t.statusname, t.client_name].filter(Boolean).join(" · "),
                isSubtle: true,
                size: "Small",
                spacing: "Small",
              },
            ]
          : []),
      ],
      actions: [
        {
          type: "Action.OpenUrl",
          title: "Open in HaloPSA",
          url,
        },
      ],
    };
    submitToTeams({
      kind: "ticket",
      card,
      html: `<a href="${url}">HaloPSA #${t.id} — ${escapeHtml(t.summary)}</a>`,
      url,
      ticketId: t.id,
    });
  };

  return (
    <>
      <div className={styles.searchRow}>
        <Field style={{ flex: 1 }}>
          <Input
            value={query}
            onChange={(_, d) => setQuery(d.value)}
            placeholder="Search ticket title, ID, or details"
            onKeyDown={(e) => {
              if (e.key === "Enter") void run();
            }}
          />
        </Field>
        <Button appearance="primary" onClick={() => void run()} disabled={loading}>
          {loading ? <Spinner size="extra-tiny" /> : "Search"}
        </Button>
      </div>
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      <div className={styles.results}>
        {results.length === 0 && !loading && (
          <Text className={styles.empty}>Search to find tickets to link.</Text>
        )}
        {results.map((t) => (
          <div key={t.id} className={styles.row} onClick={() => insert(t)}>
            <Text weight="semibold">
              #{t.id} · {t.summary}
            </Text>
            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
              {[t.statusname, t.client_name].filter(Boolean).join(" · ") || "—"}
            </Text>
          </div>
        ))}
      </div>
      <div className={styles.footer}>
        <Button appearance="subtle" onClick={() => submitToTeams(undefined)}>
          Cancel
        </Button>
      </div>
    </>
  );
}

// ---------- Canned text picker ----------

function CannedPicker({ styles }: { styles: ReturnType<typeof useStyles> }) {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<HaloCannedText[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();

  // Canned-text search runs against an in-memory cache (see halo-api.ts), so
  // it's cheap to re-run on every keystroke once the cache is warm.
  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(undefined);
    searchCannedText(query)
      .then((r) => !cancelled && setResults(r))
      .catch((e) => !cancelled && setError((e as Error).message))
      .finally(() => !cancelled && setLoading(false));
    return () => {
      cancelled = true;
    };
  }, [query]);

  const insert = (c: HaloCannedText) => {
    const html = c.html ?? `<p>${escapeHtml(c.text ?? "")}</p>`;
    submitToTeams({
      kind: "canned",
      html,
      text: c.text ?? "",
      name: c.name,
    });
  };

  return (
    <>
      <Field>
        <Input
          value={query}
          onChange={(_, d) => setQuery(d.value)}
          placeholder="Filter by name or body"
        />
      </Field>
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      <div className={styles.results}>
        {loading && (
          <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
            <Spinner size="extra-tiny" /> <Text size={200}>Loading canned text…</Text>
          </div>
        )}
        {!loading && results.length === 0 && (
          <Text className={styles.empty}>No matches.</Text>
        )}
        {results.map((c) => (
          <div key={c.id} className={styles.row} onClick={() => insert(c)}>
            <Text weight="semibold">{c.name}</Text>
            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
              {previewOf(c)}
            </Text>
          </div>
        ))}
      </div>
      <div className={styles.footer}>
        <Button appearance="subtle" onClick={() => submitToTeams(undefined)}>
          Cancel
        </Button>
      </div>
    </>
  );
}

function previewOf(c: HaloCannedText): string {
  const raw = c.text ?? c.html ?? "";
  return raw
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Hand the picked payload back to the Teams host. Teams forwards this to the
 * registered message-extension action handler, which inserts the content into
 * the chat compose box.
 *
 * Passing `undefined` cancels the dialog without an insertion.
 */
function submitToTeams(payload: unknown): void {
  try {
    // The dialog.url.submit signature is (result?, appIds?). Older SDK versions
    // expose dialog.submit on the root namespace; the new one is
    // dialog.url.submit. We try the new one first.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const d = dialog as any;
    if (d.url?.submit) {
      d.url.submit(payload);
    } else if (d.submit) {
      d.submit(payload);
    } else {
      // Dev fallback: log it so it's at least visible.
      console.log("Dialog submit (no Teams host):", payload);
    }
  } catch (e) {
    console.error("Failed to submit dialog result:", e);
  }
}
