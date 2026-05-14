import { useEffect, useState, useCallback, lazy, Suspense } from "react";
import { makeStyles, tokens, Spinner, Text } from "@fluentui/react-components";
import { ConfigScreen } from "./components/ConfigScreen";
import { AuthScreen } from "./components/AuthScreen";
import { getConfig } from "./lib/config";
import { isAuthenticated } from "./lib/auth";
import { getCurrentEmailContext, type EmailContext } from "./lib/office";

// Dashboard pulls in the bulk of the UI surface (pickers, dialogs, ticket list, log actions).
// Lazy-load it so first-launch config/auth flows ship a smaller bundle.
const Dashboard = lazy(() =>
  import("./components/Dashboard").then((m) => ({ default: m.Dashboard })),
);

type Phase = "loading" | "needs-config" | "needs-auth" | "ready";

const useStyles = makeStyles({
  root: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  center: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "12px",
    padding: "24px",
  },
});

export function App() {
  const styles = useStyles();
  const [phase, setPhase] = useState<Phase>("loading");
  const [email, setEmail] = useState<EmailContext | undefined>(undefined);

  // Compute the current phase from storage state
  const refreshPhase = useCallback(() => {
    if (!getConfig()) setPhase("needs-config");
    else if (!isAuthenticated()) setPhase("needs-auth");
    else setPhase("ready");
  }, []);

  // Re-read current email when Outlook selection changes
  const refreshEmail = useCallback(() => {
    getCurrentEmailContext()
      .then((ctx) => setEmail(ctx))
      .catch(() => setEmail(undefined));
  }, []);

  useEffect(() => {
    refreshPhase();
    refreshEmail();

    // Subscribe to item selection changes (user switches between emails)
    const handler = () => refreshEmail();
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      handler,
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          // Non-fatal — just log; UI still works via initial read
          console.warn("Failed to subscribe to ItemChanged:", result.error?.message);
        }
      },
    );

    return () => {
      // removeHandlerAsync removes all handlers for the event type
      Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged);
    };
  }, [refreshPhase, refreshEmail]);

  if (phase === "loading") {
    return (
      <div className={styles.root}>
        <div className={styles.center}>
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

  // ready
  if (!email) {
    return (
      <div className={styles.root}>
        <div className={styles.center}>
          <Text>Open a message to see Halo context.</Text>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <Suspense
        fallback={
          <div className={styles.center}>
            <Spinner size="small" />
          </div>
        }
      >
        <Dashboard email={email} onSignedOut={refreshPhase} />
      </Suspense>
    </div>
  );
}
