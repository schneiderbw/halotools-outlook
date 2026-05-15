import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { OptionsApp } from "./OptionsApp";

const root = document.getElementById("root");
if (!root) throw new Error("options: #root not found");

createRoot(root).render(
  <StrictMode>
    <FluentProvider theme={webLightTheme}>
      <OptionsApp />
    </FluentProvider>
  </StrictMode>,
);
