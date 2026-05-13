import { useState } from "react";
import {
  Button,
  Text,
  makeStyles,
  tokens,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import { signIn } from "../lib/auth";
import { clearConfig, getConfig } from "../lib/config";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "20px",
    height: "100%",
    alignItems: "stretch",
    justifyContent: "center",
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    textAlign: "center",
  },
  subtitle: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    textAlign: "center",
  },
  buttons: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
});

interface Props {
  onAuthenticated: () => void;
  onReconfigure: () => void;
}

export function AuthScreen({ onAuthenticated, onReconfigure }: Props) {
  const styles = useStyles();
  const [error, setError] = useState<string | undefined>();
  const [busy, setBusy] = useState(false);
  const cfg = getConfig();

  const handleSignIn = async () => {
    setError(undefined);
    setBusy(true);
    try {
      await signIn();
      onAuthenticated();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const handleReset = async () => {
    await clearConfig();
    onReconfigure();
  };

  return (
    <div className={styles.root}>
      <Text className={styles.title}>Sign in to HaloPSA</Text>
      <Text className={styles.subtitle}>
        Connected to <strong>{cfg?.haloBaseUrl}</strong>
      </Text>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.buttons}>
        <Button appearance="primary" onClick={handleSignIn} disabled={busy}>
          {busy ? "Opening sign-in…" : "Sign in with HaloPSA"}
        </Button>
        <Button appearance="subtle" onClick={handleReset} disabled={busy}>
          Use a different HaloPSA tenant
        </Button>
      </div>
    </div>
  );
}
