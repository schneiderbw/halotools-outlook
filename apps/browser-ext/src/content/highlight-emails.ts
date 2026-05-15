// Content script — runs on every (non-excluded) page.
//
// Pipeline:
//   1. on document idle + on DOM mutations (debounced), walk visible text
//      and `mailto:` links and harvest unique email addresses
//   2. ask the background SW to look each one up (cached for 10 min)
//   3. for known contacts, inject a small inline badge next to every
//      occurrence and link it to /customer?userid=N in Halo
//
// Constraints:
//   - PLAIN TS, no React, no Fluent UI — this ships on every page
//     the user visits. The bundle is the heavy cost; aim for tiny.
//   - Never modify form inputs, code/pre blocks, or contenteditables.
//   - Skip if the page is one of the known-noisy webmail origins; the
//     manifest already excludes them but a runtime guard is cheap.

import type { Request, Response } from "../lib/messages";

interface HitResponse {
  ok: true;
  hit: { user: { id: number; name: string; client_name?: string } } | null;
}

const BADGE_CLASS = "halo-ext-badge";
const PROCESSED_ATTR = "data-halo-ext-processed";
const EMAIL_RE = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;

// Local cache so repeated occurrences of the same address on one page
// only hit the SW once. The SW has its own 10-min cache.
const lookups = new Map<string, Promise<HitResponse["hit"]>>();

// Origins handled by the manifest's exclude_matches; double-guard here
// in case a future config change opens them up by accident.
const NOISY_HOSTS = new Set([
  "mail.google.com",
  "outlook.live.com",
  "outlook.office.com",
  "outlook.office365.com",
]);

if (!NOISY_HOSTS.has(location.host)) {
  init();
}

function init() {
  injectStyles();
  scan();
  // Re-scan on debounced DOM mutations — covers SPAs and infinite scroll.
  const debounced = debounce(scan, 600);
  const obs = new MutationObserver(debounced);
  obs.observe(document.body, { childList: true, subtree: true });
}

function injectStyles() {
  const css = `
    .${BADGE_CLASS} {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 14px;
      height: 14px;
      margin-left: 4px;
      border-radius: 50%;
      background: #c8102e;
      color: #fff;
      font-size: 9px;
      font-weight: 700;
      line-height: 1;
      font-family: system-ui, sans-serif;
      vertical-align: middle;
      text-decoration: none;
      cursor: pointer;
      user-select: none;
    }
    .${BADGE_CLASS}:hover {
      background: #a30c25;
    }
  `;
  const style = document.createElement("style");
  style.textContent = css;
  document.documentElement.appendChild(style);
}

function scan() {
  for (const a of document.querySelectorAll<HTMLAnchorElement>(
    `a[href^="mailto:"]:not([${PROCESSED_ATTR}])`,
  )) {
    const email = a.href.replace(/^mailto:/i, "").split("?")[0];
    if (email) {
      a.setAttribute(PROCESSED_ATTR, "1");
      void maybeAttachBadge(a, email);
    }
  }

  // Walk visible text nodes for plain email addresses. TreeWalker is
  // the cheapest way to do this without parsing the world.
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      if (!node.nodeValue || !node.nodeValue.includes("@")) return NodeFilter.FILTER_REJECT;
      const parent = node.parentElement;
      if (!parent) return NodeFilter.FILTER_REJECT;
      if (parent.closest(`.${BADGE_CLASS}, script, style, code, pre, textarea, input`)) {
        return NodeFilter.FILTER_REJECT;
      }
      if (parent.isContentEditable) return NodeFilter.FILTER_REJECT;
      if (parent.hasAttribute(PROCESSED_ATTR)) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });

  const targets: { node: Text; matches: RegExpExecArray[] }[] = [];
  let n: Node | null;
  // eslint-disable-next-line no-cond-assign
  while ((n = walker.nextNode())) {
    const text = n.nodeValue ?? "";
    EMAIL_RE.lastIndex = 0;
    const matches: RegExpExecArray[] = [];
    let m: RegExpExecArray | null;
    while ((m = EMAIL_RE.exec(text)) !== null) matches.push(m);
    if (matches.length > 0) targets.push({ node: n as Text, matches });
  }

  for (const { node, matches } of targets) {
    // Mark first so we don't re-process if the mutation observer fires
    // mid-stream. We modify the DOM below which will replace `node`.
    const parent = node.parentElement;
    if (!parent) continue;
    parent.setAttribute(PROCESSED_ATTR, "1");

    // Build a new fragment, interleaving plain text and badge-decorated
    // email runs. We don't modify the text itself — the original email
    // stays as plain text; the badge is appended after it.
    const frag = document.createDocumentFragment();
    let cursor = 0;
    const text = node.nodeValue ?? "";
    for (const match of matches) {
      const idx = match.index;
      if (idx > cursor) frag.appendChild(document.createTextNode(text.slice(cursor, idx)));
      const email = match[0];
      const emailNode = document.createTextNode(email);
      frag.appendChild(emailNode);
      const placeholder = document.createElement("span");
      frag.appendChild(placeholder);
      void maybeAttachBadge(placeholder, email);
      cursor = idx + email.length;
    }
    if (cursor < text.length) frag.appendChild(document.createTextNode(text.slice(cursor)));
    node.replaceWith(frag);
  }
}

async function maybeAttachBadge(anchor: Element, email: string): Promise<void> {
  const hit = await lookup(email);
  if (!hit) return;
  const badge = document.createElement("a");
  badge.className = BADGE_CLASS;
  badge.textContent = "H";
  badge.title = `HaloPSA: ${hit.user.name}${hit.user.client_name ? " — " + hit.user.client_name : ""}`;
  badge.href = "#";
  badge.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    const req: Request = {
      kind: "openInHalo",
      path: `/customer?userid=${hit.user.id}`,
    };
    void chrome.runtime.sendMessage(req);
  });
  // For mailto anchors, append inside the link's parent so the badge
  // doesn't become part of the clickable area of the mailto: itself.
  if (anchor.tagName === "A" && anchor.parentNode) {
    anchor.parentNode.insertBefore(badge, anchor.nextSibling);
  } else {
    anchor.appendChild(badge);
  }
}

function lookup(email: string): Promise<HitResponse["hit"]> {
  const norm = email.toLowerCase();
  let pending = lookups.get(norm);
  if (pending) return pending;
  pending = (async () => {
    try {
      const req: Request = { kind: "lookupEmail", email: norm };
      const r = (await chrome.runtime.sendMessage(req)) as Response<typeof req>;
      if (r.ok) return r.hit;
      return null;
    } catch {
      // SW might not be alive yet on the very first page load; swallow.
      return null;
    }
  })();
  lookups.set(norm, pending);
  return pending;
}

function debounce<T extends (...args: never[]) => void>(fn: T, ms: number): T {
  let t: number | undefined;
  return ((...args: never[]) => {
    if (t !== undefined) clearTimeout(t);
    t = setTimeout(() => fn(...args), ms) as unknown as number;
  }) as T;
}
