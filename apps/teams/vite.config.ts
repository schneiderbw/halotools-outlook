import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

// Deployed at https://tools.iusehalo.com/teams/ — base path matches.
// Override with VITE_BASE=/ for local dev or alternative paths.
const BASE = process.env.VITE_BASE ?? "/teams/";

export default defineConfig({
  base: BASE,
  plugins: [react()],
  build: {
    outDir: "dist",
    sourcemap: true,
    rollupOptions: {
      input: {
        // Personal/channel tab bundle — what an agent sees when they pin the
        // Halo app inside Teams.
        tab: resolve(__dirname, "index.html"),
        // Message-extension picker bundle — opens in a Teams task module when
        // the agent invokes "Insert ticket link" or "Insert canned text" from
        // the compose box.
        messageExtension: resolve(__dirname, "messageExtension.html"),
        // OAuth callback page — separate entry, runs inside the popup opened by
        // microsoftTeams.authentication.authenticate(). Emits dist/auth/callback.html.
        callback: resolve(__dirname, "auth/callback.html"),
      },
    },
  },
  server: {
    port: 3001,
    // Teams sideload requires HTTPS even for localhost — use a tunnel
    // (ngrok / Cloudflare) or local certs.
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
});
