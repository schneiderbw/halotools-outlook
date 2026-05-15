import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { crx } from "@crxjs/vite-plugin";
import manifest from "./manifest.json" with { type: "json" };

// @crxjs/vite-plugin reads `manifest.json` as the source of truth and rewrites
// asset paths, splits the service worker bundle, etc. We keep the manifest as a
// plain JSON file (not a JS object) so that future Firefox compatibility — which
// will likely require swapping in a manifest with `browser_specific_settings` —
// is just a different file rather than conditional code.
export default defineConfig({
  plugins: [react(), crx({ manifest })],
  build: {
    outDir: "dist",
    emptyOutDir: true,
    target: "es2022",
    // Source maps help when debugging the popup / service worker in DevTools.
    // They add some bulk but the bundle is small and only the dev / unpacked
    // load uses them — Chrome Web Store strips them or ignores them anyway.
    sourcemap: true,
  },
  // crx plugin handles the HTML entry points declared in manifest.json,
  // so we don't need to enumerate inputs in rollupOptions.
  server: {
    port: 5173,
    strictPort: true,
    // crxjs uses HMR over a fixed port for the content scripts.
    hmr: {
      port: 5174,
    },
  },
});
