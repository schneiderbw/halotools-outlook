import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { PopupApp } from "./PopupApp";

const root = document.getElementById("root");
if (!root) throw new Error("popup: #root not found");

createRoot(root).render(
  <StrictMode>
    <FluentProvider theme={webLightTheme}>
      <PopupApp />
    </FluentProvider>
  </StrictMode>,
);
