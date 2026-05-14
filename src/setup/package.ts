import { zipSync, strToU8 } from "fflate";

export interface TenantInput {
  haloBaseUrl: string;
  clientId: string;
  appName?: string;
}

interface ManifestIcon {
  size: number;
  url: string;
}

interface ManifestExtensionGroup {
  id: string;
  label: string;
  icons: ManifestIcon[];
  controls: Array<{
    id: string;
    type: string;
    label: string;
    icons: ManifestIcon[];
    supertip: { title: string; description: string };
    actionId: string;
  }>;
}

interface ManifestRibbon {
  contexts: string[];
  tabs: Array<{
    builtInTabId: string;
    groups: ManifestExtensionGroup[];
  }>;
}

interface Manifest {
  id: string;
  name: { short: string; full: string };
  validDomains?: string[];
  icons: { outline: string; color: string };
  extensions: Array<{
    runtimes: Array<{ id: string; code: { page: string } }>;
    ribbons?: ManifestRibbon[];
  }>;
  [k: string]: unknown;
}

function hostnameOf(url: string): string {
  return new URL(url).hostname;
}

/**
 * Build a per-tenant manifest by taking the generic published manifest
 * and overlaying tenant-specific fields:
 *  - Fresh GUID so M365 treats it as a distinct app per MSP
 *  - validDomains extended with the MSP's Halo host
 *  - Runtime page URL carries ?halo=...&clientId=... so the SPA self-configures
 *    on first launch and skips the in-app config screen
 */
export function buildManifest(template: Manifest, input: TenantInput): Manifest {
  const haloHost = hostnameOf(input.haloBaseUrl);
  const params = new URLSearchParams({
    halo: input.haloBaseUrl,
    clientId: input.clientId,
  }).toString();

  const cloned: Manifest = JSON.parse(JSON.stringify(template));
  cloned.id = crypto.randomUUID();

  const validDomains = new Set(cloned.validDomains ?? []);
  validDomains.add(haloHost);
  cloned.validDomains = Array.from(validDomains);

  if (input.appName) {
    cloned.name = {
      short: input.appName.slice(0, 30),
      full: input.appName,
    };
  }

  for (const ext of cloned.extensions ?? []) {
    for (const rt of ext.runtimes ?? []) {
      const u = new URL(rt.code.page);
      // Strip /index.html if present so the param URL is clean
      if (u.pathname.endsWith("/index.html")) {
        u.pathname = u.pathname.slice(0, -"index.html".length);
      }
      u.search = params;
      rt.code.page = u.toString();
    }
  }

  return cloned;
}

async function fetchBytes(url: string): Promise<Uint8Array> {
  const r = await fetch(url, { cache: "no-store" });
  if (!r.ok) throw new Error(`Failed to fetch ${url}: ${r.status}`);
  return new Uint8Array(await r.arrayBuffer());
}

/**
 * Produce the M365 admin upload zip. Contains the customized manifest at the
 * root plus the icon files the manifest references via relative paths.
 */
export async function buildPackageZip(
  template: Manifest,
  input: TenantInput,
): Promise<Blob> {
  const manifest = buildManifest(template, input);
  const iconPaths = new Set<string>();
  if (manifest.icons?.outline) iconPaths.add(manifest.icons.outline);
  if (manifest.icons?.color) iconPaths.add(manifest.icons.color);

  const files: Record<string, Uint8Array> = {
    "manifest.json": strToU8(JSON.stringify(manifest, null, 2)),
  };

  // Icon paths in the manifest are relative to /outlook/manifest.json, not to
  // the wizard's own URL — anchor against the manifest's location explicitly.
  const iconBase = `${window.location.origin}/outlook/`;
  await Promise.all(
    Array.from(iconPaths).map(async (rel) => {
      const url = new URL(rel, iconBase).toString();
      files[rel] = await fetchBytes(url);
    }),
  );

  const zipped = zipSync(files, { level: 6 });
  return new Blob([zipped.buffer as ArrayBuffer], { type: "application/zip" });
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export async function fetchTemplate(url: string): Promise<Manifest> {
  const r = await fetch(url, { cache: "no-store" });
  if (!r.ok) throw new Error(`Failed to fetch manifest template: ${r.status}`);
  return (await r.json()) as Manifest;
}
