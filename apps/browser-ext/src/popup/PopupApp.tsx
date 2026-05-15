import { useEffect, useState } from "react";
import {
  Button,
  Input,
  Spinner,
  Text,
  Title3,
  Caption1,
  Body1,
  Divider,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  PersonRegular,
  TicketDiagonalRegular,
  OpenRegular,
  AddRegular,
} from "@fluentui/react-icons";
import { send } from "../lib/messages";
import type { HaloUser } from "../lib/types";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    minHeight: "160px",
  },
  centered: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    alignItems: "center",
    justifyContent: "center",
    padding: tokens.spacingVerticalXL,
    textAlign: "center",
  },
  searchRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
  },
  resultCard: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  resultHeader: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  actions: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalXS,
  },
});

type Status =
  | { kind: "loading" }
  | { kind: "needsConfig" }
  | { kind: "needsSignIn" }
  | { kind: "ready" };

export function PopupApp() {
  const styles = useStyles();
  const [status, setStatus] = useState<Status>({ kind: "loading" });
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<HaloUser[]>([]);
  const [searching, setSearching] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    void refreshStatus();
  }, []);

  async function refreshStatus() {
    const r = await send({ kind: "getAuthStatus" });
    if (!r.ok) {
      setStatus({ kind: "needsConfig" });
      return;
    }
    if (!r.configured) setStatus({ kind: "needsConfig" });
    else if (!r.signedIn) setStatus({ kind: "needsSignIn" });
    else setStatus({ kind: "ready" });
  }

  async function runSearch(q: string) {
    setError(null);
    if (!q.trim()) {
      setResults([]);
      return;
    }
    setSearching(true);
    try {
      const r = await send({ kind: "search", query: q.trim() });
      if (r.ok) setResults(r.users);
      else setError(r.error);
    } finally {
      setSearching(false);
    }
  }

  function openOptions() {
    chrome.runtime.openOptionsPage();
  }

  async function signInFromPopup() {
    // chrome.identity.launchWebAuthFlow used to be allowed from popup
    // contexts, but the popup closes the moment focus moves to the
    // auth window — which kills the in-flight promise. Sending the
    // message to the SW means the SW owns the flow and the popup can
    // close. We poll auth status when it reopens.
    setStatus({ kind: "loading" });
    const r = await send({ kind: "signIn" });
    if (!r.ok) {
      setError(r.error);
    }
    await refreshStatus();
  }

  if (status.kind === "loading") {
    return (
      <div className={styles.centered}>
        <Spinner size="small" label="Loading" />
      </div>
    );
  }

  if (status.kind === "needsConfig") {
    return (
      <div className={styles.centered}>
        <Title3>HaloPSA Tools</Title3>
        <Caption1>Connect to your Halo tenant to get started.</Caption1>
        <Button appearance="primary" onClick={openOptions}>
          Set up Halo
        </Button>
      </div>
    );
  }

  if (status.kind === "needsSignIn") {
    return (
      <div className={styles.centered}>
        <Title3>Sign in to Halo</Title3>
        <Caption1>Your tenant is configured but you're not signed in.</Caption1>
        <Button appearance="primary" onClick={signInFromPopup}>
          Sign in
        </Button>
        {error && <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{error}</Text>}
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <Title3>HaloPSA</Title3>
      <div className={styles.searchRow}>
        <Input
          autoFocus
          placeholder="Email or name"
          value={query}
          onChange={(_, d) => {
            setQuery(d.value);
            void runSearch(d.value);
          }}
          style={{ flex: 1 }}
        />
      </div>
      {searching && <Spinner size="tiny" label="Searching" />}
      {error && <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{error}</Text>}
      {!searching && results.length === 0 && query.trim() && (
        <Caption1>No matches.</Caption1>
      )}
      {results.map((u) => (
        <UserCard key={u.id} user={u} />
      ))}
      {results.length > 0 && <Divider />}
      <Button
        icon={<AddRegular />}
        appearance="subtle"
        onClick={() => send({ kind: "openInHalo", path: "/tickets?view=new" })}
      >
        New ticket
      </Button>
    </div>
  );
}

function UserCard({ user }: { user: HaloUser }) {
  const styles = useStyles();
  return (
    <div className={styles.resultCard}>
      <div className={styles.resultHeader}>
        <PersonRegular />
        <Body1>
          <strong>{user.name}</strong>
        </Body1>
      </div>
      <Caption1>{user.emailaddress}</Caption1>
      {user.client_name && <Caption1>{user.client_name}</Caption1>}
      <div className={styles.actions}>
        <Button
          icon={<OpenRegular />}
          size="small"
          onClick={() =>
            send({ kind: "openInHalo", path: `/customer?userid=${user.id}` })
          }
        >
          Open
        </Button>
        <Button
          icon={<TicketDiagonalRegular />}
          size="small"
          onClick={() =>
            send({
              kind: "openInHalo",
              path: `/tickets?view=new&user_id=${user.id}`,
            })
          }
        >
          New ticket
        </Button>
      </div>
    </div>
  );
}
