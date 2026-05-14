import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

// Deployed at https://tools.iusehalo.com/outlook/ — base path matches.
// Override with VITE_BASE=/ for local dev or alternative paths.
const BASE = process.env.VITE_BASE ?? "/outlook/";

export default defineConfig({
  base: BASE,
  plugins: [react()],
  build: {
    outDir: "dist",
    sourcemap: true,
    rollupOptions: {
      input: {
        // Main task pane bundle
        taskpane: resolve(__dirname, "index.html"),
        // OAuth callback page — separate entry so it can run standalone in the Office Dialog.
        // Lives at <root>/auth/callback.html so Vite emits dist/auth/callback.html.
        callback: resolve(__dirname, "auth/callback.html"),
        // Compose-surface task pane — mounted in the mailCompose runtime declared in manifest.json.
        // Emits dist/compose/index.html, served at /outlook/compose/.
        compose: resolve(__dirname, "compose/index.html"),
        // Admin setup wizard — generates per-tenant manifest .zip for M365 upload.
        // Standalone page (no Office.js); served at /outlook/setup/.
        setup: resolve(__dirname, "setup/index.html"),
      },
    },
  },
  server: {
    port: 3000,
    // Outlook sideload requires HTTPS — use office-addin-dev-certs locally,
    // or tunnel via ngrok / Cloudflare. Production hosts handle TLS.
    headers: {
      // Loosen for local dev only; production CORS is owned by Halo per-tenant.
      "Access-Control-Allow-Origin": "*",
    },
  },
});
