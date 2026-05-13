import { useState } from "react";
import {
  Button,
  Field,
  Input,
  Text,
  Title2,
  Subtitle2,
  Body1,
  Body1Strong,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  makeStyles,
  tokens,
  Card,
  CardHeader,
} from "@fluentui/react-components";
import {
  CheckmarkCircle24Filled,
  Copy24Regular,
  ArrowDownload24Regular,
} from "@fluentui/react-icons";
import {
  buildPackageZip,
  downloadBlob,
  fetchTemplate,
} from "./package";

type Step = "halo-url" | "register-app" | "client-id" | "done";

const REDIRECT_URI = `${window.location.origin}/outlook/auth/callback.html`;
const MANIFEST_TEMPLATE_URL = "/outlook/manifest.json";

const useStyles = makeStyles({
  page: {
    minHeight: "100vh",
    display: "flex",
    justifyContent: "center",
    padding: "32px 16px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  card: {
    width: "min(640px, 100%)",
    padding: "28px",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  header: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  helpText: {
    color: tokens.colorNeutralForeground3,
  },
  steps: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  stepActive: {
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
  codeRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "8px 12px",
    borderRadius: tokens.borderRadiusMedium,
    fontFamily: "Consolas, monospace",
    fontSize: tokens.fontSizeBase200,
    wordBreak: "break-all",
  },
  buttonRow: {
    display: "flex",
    gap: "8px",
    justifyContent: "flex-end",
    marginTop: "8px",
  },
  successBody: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  ol: {
    margin: 0,
    paddingLeft: "20px",
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
});

function CopyableValue({ value }: { value: string }) {
  const styles = useStyles();
  const [copied, setCopied] = useState(false);
  const copy = async () => {
    await navigator.clipboard.writeText(value);
    setCopied(true);
    setTimeout(() => setCopied(false), 1500);
  };
  return (
    <div className={styles.codeRow}>
      <span style={{ flex: 1 }}>{value}</span>
      <Button
        size="small"
        appearance="subtle"
        icon={copied ? <CheckmarkCircle24Filled /> : <Copy24Regular />}
        onClick={copy}
        aria-label="Copy"
      >
        {copied ? "Copied" : "Copy"}
      </Button>
    </div>
  );
}

function StepIndicator({ step }: { step: Step }) {
  const styles = useStyles();
  const order: Step[] = ["halo-url", "register-app", "client-id", "done"];
  const idx = order.indexOf(step);
  return (
    <div className={styles.steps}>
      {order.slice(0, 3).map((s, i) => (
        <span key={s} className={idx === i ? styles.stepActive : undefined}>
          {i + 1}. {s === "halo-url" ? "Halo URL" : s === "register-app" ? "Register app" : "Client ID"}
          {i < 2 ? " →" : ""}
        </span>
      ))}
    </div>
  );
}

export function SetupApp() {
  const styles = useStyles();
  const [step, setStep] = useState<Step>("halo-url");
  const [haloUrl, setHaloUrl] = useState("");
  const [clientId, setClientId] = useState("");
  const [error, setError] = useState<string | undefined>();
  const [building, setBuilding] = useState(false);

  const normalizedHalo = haloUrl.trim().replace(/\/+$/, "");
  const haloUrlValid = /^https:\/\/[^\s]+\.[^\s]+$/.test(normalizedHalo);

  const submitHaloUrl = () => {
    setError(undefined);
    if (!haloUrlValid) {
      setError("Enter a valid HaloPSA URL starting with https://");
      return;
    }
    setStep("register-app");
  };

  const downloadPackage = async () => {
    setError(undefined);
    setBuilding(true);
    try {
      const template = await fetchTemplate(MANIFEST_TEMPLATE_URL);
      const zip = await buildPackageZip(template, {
        haloBaseUrl: normalizedHalo,
        clientId: clientId.trim(),
      });
      const slug = new URL(normalizedHalo).hostname.replace(/\./g, "-");
      downloadBlob(zip, `halo-outlook-${slug}.zip`);
      setStep("done");
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBuilding(false);
    }
  };

  return (
    <div className={styles.page}>
      <Card className={styles.card}>
        <CardHeader
          header={
            <div className={styles.header}>
              <Title2>Halo for Outlook — setup</Title2>
              <Body1 className={styles.helpText}>
                Generate a tenant-specific Outlook app package for your MSP. Upload
                the resulting zip in the Microsoft 365 admin center and your team
                gets the add-in preconfigured for your HaloPSA instance.
              </Body1>
            </div>
          }
        />

        {step !== "done" && <StepIndicator step={step} />}

        {step === "halo-url" && (
          <>
            <Field
              label="HaloPSA URL"
              required
              hint="The base URL of your Halo instance, e.g. https://halo.yourcompany.com"
            >
              <Input
                value={haloUrl}
                onChange={(_, d) => setHaloUrl(d.value)}
                placeholder="https://halo.yourcompany.com"
                autoComplete="off"
              />
            </Field>
            {error && (
              <MessageBar intent="error">
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}
            <div className={styles.buttonRow}>
              <Button
                appearance="primary"
                disabled={!haloUrl.trim()}
                onClick={submitHaloUrl}
              >
                Next
              </Button>
            </div>
          </>
        )}

        {step === "register-app" && (
          <>
            <Subtitle2>Register a Halo Connect application</Subtitle2>
            <Body1>In your Halo instance, go to:</Body1>
            <Body1Strong>Configuration → Integrations → Halo Connect → API</Body1Strong>
            <Body1>Click <em>New</em> and configure the application as follows:</Body1>
            <ol className={styles.ol}>
              <li><Body1>Application Name: <code>Outlook Add-in</code> (anything you like)</Body1></li>
              <li><Body1>Authentication Method: <code>Authorization Code (PKCE)</code></Body1></li>
              <li><Body1>Login Type: <code>Agent</code></Body1></li>
              <li>
                <Body1>Redirect URI — paste exactly:</Body1>
                <CopyableValue value={REDIRECT_URI} />
              </li>
              <li>
                <Body1>Permissions / Scopes: <code>all</code> (or the minimum needed for tickets, users, clients, actions, attachments)</Body1>
              </li>
              <li>
                <Body1>Under the application's CORS whitelist, add:</Body1>
                <CopyableValue value={window.location.origin} />
              </li>
              <li><Body1>Save the application, then copy the generated <strong>Client ID</strong>.</Body1></li>
            </ol>
            <div className={styles.buttonRow}>
              <Button appearance="secondary" onClick={() => setStep("halo-url")}>Back</Button>
              <Button appearance="primary" onClick={() => setStep("client-id")}>I've registered the app</Button>
            </div>
          </>
        )}

        {step === "client-id" && (
          <>
            <Field
              label="Client ID"
              required
              hint="From the Halo Connect application you just registered"
            >
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
            <div className={styles.buttonRow}>
              <Button appearance="secondary" onClick={() => setStep("register-app")} disabled={building}>Back</Button>
              <Button
                appearance="primary"
                icon={<ArrowDownload24Regular />}
                disabled={!clientId.trim() || building}
                onClick={downloadPackage}
              >
                {building ? "Building…" : "Download package"}
              </Button>
            </div>
          </>
        )}

        {step === "done" && (
          <>
            <MessageBar intent="success">
              <MessageBarBody>
                <MessageBarTitle>Package ready</MessageBarTitle>
                Now upload it to Microsoft 365.
              </MessageBarBody>
            </MessageBar>
            <div className={styles.successBody}>
              <Subtitle2>Next: deploy to your tenant</Subtitle2>
              <ol className={styles.ol}>
                <li><Body1>Sign in at <strong>admin.microsoft.com</strong> as a global or apps admin.</Body1></li>
                <li><Body1>Go to <strong>Settings → Integrated apps</strong>.</Body1></li>
                <li><Body1>Click <strong>Upload custom apps</strong>.</Body1></li>
                <li><Body1>Choose <strong>App type: Office Add-in</strong> and <strong>Upload manifest file (.zip)</strong>, then pick the file you just downloaded.</Body1></li>
                <li><Body1>Assign to your team and finish the deployment.</Body1></li>
              </ol>
              <Text className={styles.helpText}>
                Microsoft pushes the add-in to assigned users within a few hours. On first launch
                it auto-configures for your HaloPSA instance — no per-user setup.
              </Text>
              <div className={styles.buttonRow}>
                <Button onClick={() => { setStep("halo-url"); setHaloUrl(""); setClientId(""); }}>
                  Build another package
                </Button>
              </div>
            </div>
          </>
        )}
      </Card>
    </div>
  );
}
