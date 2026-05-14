import { useRef, useState } from "react";
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
  readExistingManifest,
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
  // When set, the regenerated package keeps this GUID so M365 admin treats
  // the upload as an update to the existing deployment instead of a brand-new
  // app. Populated by readExistingManifest after an admin uploads their
  // current package.
  const [existingAppId, setExistingAppId] = useState<string | undefined>();
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

  const handleExistingPackage = async (file: File) => {
    setError(undefined);
    try {
      const extracted = await readExistingManifest(file);
      setExistingAppId(extracted.id);
      if (extracted.haloBaseUrl) setHaloUrl(extracted.haloBaseUrl);
      if (extracted.clientId) setClientId(extracted.clientId);
      // Both fields pulled — jump straight to the final step so the admin
      // just hits Download.
      if (extracted.haloBaseUrl && extracted.clientId) {
        setStep("client-id");
      } else if (extracted.haloBaseUrl) {
        setStep("register-app");
      }
    } catch (e) {
      setError(`Couldn't read package: ${(e as Error).message}`);
    }
  };

  const downloadPackage = async () => {
    setError(undefined);
    setBuilding(true);
    try {
      const template = await fetchTemplate(MANIFEST_TEMPLATE_URL);
      const zip = await buildPackageZip(template, {
        haloBaseUrl: normalizedHalo,
        clientId: clientId.trim(),
        existingAppId,
      });
      const slug = new URL(normalizedHalo).hostname.replace(/\./g, "-");
      const filename = existingAppId
        ? `halo-outlook-${slug}-update.zip`
        : `halo-outlook-${slug}.zip`;
      downloadBlob(zip, filename);
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

            <ExistingPackageUpload onPick={handleExistingPackage} />
          </>
        )}

        {existingAppId && step !== "halo-url" && step !== "done" && (
          <MessageBar intent="info">
            <MessageBarBody>
              Update mode — reusing app ID <code>{existingAppId.slice(0, 8)}…</code>.
              The package you download will replace the existing deployment in M365 admin
              instead of installing a new app.
            </MessageBarBody>
          </MessageBar>
        )}

        {step === "register-app" && (
          <>
            <Subtitle2>Register an API application in Halo</Subtitle2>
            <Body1>In your Halo instance, go to:</Body1>
            <Body1Strong>Configuration → Integrations → API</Body1Strong>
            <Body1>Click <em>New</em> and configure the application as follows:</Body1>
            <ol className={styles.ol}>
              <li><Body1>Application Name: <code>Outlook Add-in</code> (anything you like)</Body1></li>
              <li><Body1>Authentication Method: <code>Authorization Code (PKCE)</code></Body1></li>
              <li><Body1>Allowed Logins: check <strong>Allow Agent Logins</strong>. Leave <em>Allow End-User Logins</em> and <em>Allow Anonymous Access</em> unchecked.</Body1></li>
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
              hint="From the API application you just registered in Halo"
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
                <MessageBarTitle>
                  {existingAppId ? "Update package ready" : "Package ready"}
                </MessageBarTitle>
                {existingAppId
                  ? "Upload it to Microsoft 365 — M365 will replace the existing deployment because the app ID matches."
                  : "Now upload it to Microsoft 365."}
              </MessageBarBody>
            </MessageBar>
            <div className={styles.successBody}>
              <Subtitle2>Next: deploy to your tenant</Subtitle2>
              <ol className={styles.ol}>
                <li>
                  <Body1>
                    Sign in to{" "}
                    <a
                      href="https://admin.microsoft.com/Adminportal/Home#/Settings/IntegratedApps"
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      admin.microsoft.com → Integrated apps
                    </a>{" "}
                    as a global or apps admin.
                  </Body1>
                </li>
                <li>
                  <Body1>
                    Click <strong>Upload custom apps</strong>.
                  </Body1>
                </li>
                <li>
                  <Body1>
                    <strong>App type: Microsoft Teams app</strong>{" "}
                    <em>(yes, Teams — this is the path that accepts the unified manifest zip and</em>{" "}
                    <em>also deploys to Outlook. The "Office Add-in" option only takes the legacy</em>{" "}
                    <em>XML format and won't work here).</em>
                  </Body1>
                </li>
                <li>
                  <Body1>
                    <strong>Upload an app package (.zip)</strong> → pick the file you just
                    downloaded.
                  </Body1>
                </li>
                <li>
                  <Body1>Assign to your team and finish the deployment.</Body1>
                </li>
                <li>
                  <Body1>
                    Also surface it in{" "}
                    <a
                      href="https://admin.teams.microsoft.com/policies/app-setup"
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      Teams admin → App setup policies
                    </a>{" "}
                    if you want it pre-pinned for users.
                  </Body1>
                </li>
              </ol>
              <Text className={styles.helpText}>
                Microsoft pushes the add-in to assigned users within a few hours. On first launch
                it auto-configures for your HaloPSA instance — no per-user setup.
              </Text>
              <div className={styles.buttonRow}>
                <Button
                  onClick={() => {
                    setStep("halo-url");
                    setHaloUrl("");
                    setClientId("");
                    setExistingAppId(undefined);
                  }}
                >
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

function ExistingPackageUpload({ onPick }: { onPick: (file: File) => void }) {
  const inputRef = useRef<HTMLInputElement>(null);
  return (
    <div
      style={{
        marginTop: 16,
        paddingTop: 12,
        borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
        display: "flex",
        flexDirection: "column",
        gap: 4,
      }}
    >
      <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
        Updating an existing deployment? Upload your current package to keep the
        same app ID — M365 will treat the result as an update instead of a new install.
      </Body1>
      <div>
        <Button
          appearance="subtle"
          size="small"
          onClick={() => inputRef.current?.click()}
        >
          Upload existing package…
        </Button>
        <input
          ref={inputRef}
          type="file"
          accept=".zip,.json,application/zip,application/json"
          style={{ display: "none" }}
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) onPick(file);
            // Reset so re-picking the same file fires again.
            e.target.value = "";
          }}
        />
      </div>
    </div>
  );
}
