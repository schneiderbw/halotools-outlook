import { useEffect, useState } from "react";
import {
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Link,
  makeStyles,
  tokens,
} from "@fluentui/react-components";

interface Latest {
  manifestVersion: string;
  released?: string;
}

const useStyles = makeStyles({
  wrap: {
    padding: "8px 12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

// Compare the first three segments of two Office-style versions (`major.minor.patch[.revision]`).
// Revision (4th segment) is intentionally ignored — it bumps on every wizard regeneration and
// doesn't signal a real manifest change.
function baseVersion(v: string): string {
  return v.split(".").slice(0, 3).join(".");
}

export function UpdateBanner() {
  const styles = useStyles();
  const [latest, setLatest] = useState<Latest | null>(null);
  const [installed, setInstalled] = useState<string | null>(null);

  useEffect(() => {
    // The installed manifest version is stamped into the runtime URL by the
    // setup wizard. Older installs predating the version stamp will be missing
    // `mv` entirely — surface those as outdated too, since re-uploading is
    // what gets them onto the new versioning scheme.
    const mv = new URLSearchParams(window.location.search).get("mv");
    setInstalled(mv);

    // Fetch latest.json from our origin. Resolve relative to the SPA root so
    // it works whether served at /outlook/ or /.
    fetch(new URL("latest.json", document.baseURI).toString(), { cache: "no-store" })
      .then((r) => (r.ok ? (r.json() as Promise<Latest>) : null))
      .then((data) => setLatest(data))
      .catch(() => setLatest(null));
  }, []);

  if (!latest) return null;
  if (installed && baseVersion(installed) === baseVersion(latest.manifestVersion)) {
    return null;
  }

  const setupHref = new URL("setup/", document.baseURI).toString();

  return (
    <div className={styles.wrap}>
      <MessageBar intent="warning">
        <MessageBarBody>
          <MessageBarTitle>Add-in update available</MessageBarTitle>
          {installed
            ? ` This add-in is version ${baseVersion(installed)}. Version ${latest.manifestVersion} is available.`
            : ` Version ${latest.manifestVersion} is available.`}{" "}
          Ask your M365 admin to download the updated manifest and re-upload it via the
          Microsoft 365 admin center → Integrated apps.{" "}
          <Link href={setupHref} target="_blank" rel="noopener noreferrer">
            Open setup
          </Link>
        </MessageBarBody>
      </MessageBar>
    </div>
  );
}
