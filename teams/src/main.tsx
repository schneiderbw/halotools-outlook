import React from "react";
import ReactDOM from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { app as teamsApp } from "@microsoft/teams-js";
import { TabApp } from "./tab/TabApp";

class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  { error?: Error }
> {
  state: { error?: Error } = {};
  static getDerivedStateFromError(error: Error) {
    return { error };
  }
  componentDidCatch(error: Error, info: React.ErrorInfo) {
    console.error("Halo Teams app render error:", error, info);
  }
  render() {
    if (this.state.error) {
      return (
        <div
          style={{
            padding: 16,
            fontFamily: "Segoe UI, system-ui, sans-serif",
            color: "#323130",
          }}
        >
          <div style={{ color: "#a4262c", fontWeight: 600, marginBottom: 8 }}>
            Halo Teams app crashed
          </div>
          <div
            style={{
              fontSize: 13,
              fontFamily: "Consolas, monospace",
              whiteSpace: "pre-wrap",
              background: "#faf9f8",
              border: "1px solid #edebe9",
              padding: 8,
              borderRadius: 4,
            }}
          >
            {this.state.error.message}
            {this.state.error.stack ? "\n\n" + this.state.error.stack : ""}
          </div>
          <button
            style={{ marginTop: 12, padding: "6px 12px" }}
            onClick={() => this.setState({ error: undefined })}
          >
            Try again
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

async function bootstrap() {
  // teams-js initialization is required before any other Teams API call.
  // app.initialize() resolves immediately when running in a normal browser tab
  // outside Teams — it throws after a timeout. We swallow the rejection so the
  // tab is still usable in plain-browser dev (auth then falls back to the
  // window.open / postMessage path implemented in lib/auth.ts).
  try {
    await teamsApp.initialize();
  } catch {
    /* not inside Teams — dev mode */
  }

  ReactDOM.createRoot(document.getElementById("root")!).render(
    <React.StrictMode>
      <FluentProvider theme={webLightTheme}>
        <ErrorBoundary>
          <TabApp />
        </ErrorBoundary>
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
