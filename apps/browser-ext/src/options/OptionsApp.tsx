import { useEffect, useState } from "react";
import {
  Button,
  Field,
  Input,
  Spinner,
  Text,
  Title2,
  Title3,
  Caption1,
  Body1,
  Divider,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Card,
  CardHeader,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { getConfig, setConfig, clearConfig, type TenantConfig } from "../lib/storage";
import { send } from "../lib/messages";

const useStyles = makeStyles({
  page: {
    maxWidth: "640px",
    margin: "0 auto",
    padding: tokens.spacingVerticalXXXL,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
  },
  row: {
    display: "flex",
    gap: tokens.spacingHorizontalM,
    alignItems: "flex-end",
  },
  redirect: {
    fontFamily: tokens.fontFamilyMonospace,
    backgroundColor: tokens.colorNeutralBackground3,
    padding: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusSmall,
    wordBreak: "break-all",
  },
});

type SaveState = { kind: "idle" } | { kind: "saving" } | { kind: "error"; msg: string };

export function OptionsApp() {
  const styles = useStyles();
  const [haloBaseUrl, setHaloBaseUrl] = useState("");
  const [clientId, setClientId] = useState("");
  const [scope, setScope] = useState("all");
  const [save, setSave] = useState<SaveState>({ kind: "idle" });
  const [authStatus, setAuthStatus] = useState<{
    configured: boolean;
    signedIn: boolean;
  } | null>(null);
  const [authBusy, setAuthBusy] = useState(false);
  const [authError, setAuthError] = useState<string | null>(null);
  const redirectUri = chrome.identity?.getRedirectURL?.() ?? "(unavailable)";

  useEffect(() => {
    void (async () => {
      const cfg = await getConfig();
      if (cfg) {
        setHaloBaseUrl(cfg.haloBaseUrl);
        setClientId(cfg.clientId);
        if (cfg.scope) setScope(cfg.scope);
      }
      await refreshAuth();
    })();
  }, []);

  async function refreshAuth() {
    const r = await send({ kind: "getAuthStatus" });
    if (r.ok) setAuthStatus({ configured: r.configured, signedIn: r.signedIn });
  }

  async function handleSave() {
    setSave({ kind: "saving" });
    try {
      const cfg: TenantConfig = {
        haloBaseUrl: haloBaseUrl.trim(),
        clientId: clientId.trim(),
        scope: scope.trim() || undefined,
      };
      await setConfig(cfg);
      setSave({ kind: "idle" });
      await refreshAuth();
    } catch (e) {
      setSave({ kind: "error", msg: (e as Error).message });
    }
  }

  async function handleConnect() {
    setAuthError(null);
    setAuthBusy(true);
    try {
      const r = await send({ kind: "signIn" });
      if (!r.ok) setAuthError(r.error);
    } finally {
      setAuthBusy(false);
      await refreshAuth();
    }
  }

  async function handleSignOut() {
    await send({ kind: "signOut" });
    await refreshAuth();
  }

  async function handleReset() {
    if (!confirm("Clear stored tenant config and tokens?")) return;
    await clearConfig();
    setHaloBaseUrl("");
    setClientId("");
    setScope("all");
    await refreshAuth();
  }

  return (
    <div className={styles.page}>
      <Title2>HaloPSA Tools — Options</Title2>
      <Caption1>
        Connect this extension to your HaloPSA tenant. Each MSP brings its own Halo URL
        and Halo Connect Client ID.
      </Caption1>

      <Card>
        <CardHeader header={<Title3>1. Register the redirect URI on Halo</Title3>} />
        <Body1>
          In your Halo Connect application, add this exact URL to the allowed redirect
          URIs and to the CORS Whitelist:
        </Body1>
        <div className={styles.redirect}>{redirectUri}</div>
        <Caption1>
          This URL is unique to this installation of the extension. If you reinstall or
          load it from a different location, the URL changes — update Halo accordingly.
        </Caption1>
      </Card>

      <Card>
        <CardHeader header={<Title3>2. Tenant details</Title3>} />
        <Field label="HaloPSA URL" required hint="e.g. https://halo.example.com (no trailing slash)">
          <Input
            value={haloBaseUrl}
            onChange={(_, d) => setHaloBaseUrl(d.value)}
            placeholder="https://halo.example.com"
          />
        </Field>
        <Field label="Halo Connect Client ID" required>
          <Input value={clientId} onChange={(_, d) => setClientId(d.value)} />
        </Field>
        <Field label="Scope" hint="Defaults to 'all' — leave alone unless your tenant restricts scopes.">
          <Input value={scope} onChange={(_, d) => setScope(d.value)} />
        </Field>
        {save.kind === "error" && (
          <MessageBar intent="error">
            <MessageBarBody>
              <MessageBarTitle>Couldn't save</MessageBarTitle>
              {save.msg}
            </MessageBarBody>
          </MessageBar>
        )}
        <div className={styles.row}>
          <Button
            appearance="primary"
            disabled={save.kind === "saving" || !haloBaseUrl || !clientId}
            onClick={handleSave}
          >
            {save.kind === "saving" ? <Spinner size="tiny" /> : "Save"}
          </Button>
          <Button appearance="subtle" onClick={handleReset}>
            Clear stored config
          </Button>
        </div>
      </Card>

      <Card>
        <CardHeader header={<Title3>3. Connect</Title3>} />
        {!authStatus?.configured && (
          <Caption1>Save your tenant details above, then return here to connect.</Caption1>
        )}
        {authStatus?.configured && (
          <>
            <Body1>
              Status:{" "}
              {authStatus.signedIn ? (
                <Text weight="semibold" style={{ color: tokens.colorPaletteGreenForeground1 }}>
                  Signed in
                </Text>
              ) : (
                <Text style={{ color: tokens.colorPaletteRedForeground1 }}>Not signed in</Text>
              )}
            </Body1>
            {authError && (
              <MessageBar intent="error">
                <MessageBarBody>{authError}</MessageBarBody>
              </MessageBar>
            )}
            <Divider />
            <div className={styles.row}>
              {!authStatus.signedIn ? (
                <Button appearance="primary" onClick={handleConnect} disabled={authBusy}>
                  {authBusy ? <Spinner size="tiny" /> : "Connect to Halo"}
                </Button>
              ) : (
                <Button onClick={handleSignOut}>Sign out</Button>
              )}
            </div>
          </>
        )}
      </Card>
    </div>
  );
}
