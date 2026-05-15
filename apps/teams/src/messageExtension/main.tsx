import React from "react";
import ReactDOM from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { app as teamsApp } from "@microsoft/teams-js";
import { ExtensionApp } from "./ExtensionApp";

async function bootstrap() {
  try {
    await teamsApp.initialize();
  } catch {
    /* not inside Teams — dev mode */
  }
  ReactDOM.createRoot(document.getElementById("root")!).render(
    <React.StrictMode>
      <FluentProvider theme={webLightTheme}>
        <ExtensionApp />
      </FluentProvider>
    </React.StrictMode>,
  );
}

bootstrap().catch((err) => {
  const root = document.getElementById("root");
  if (root) {
    root.innerHTML = `<div style="padding:16px;color:#a4262c;font-family:Segoe UI,system-ui,sans-serif">
      <strong>Initialization failed</strong><br/>${String(err)}
    </div>`;
  }
});
