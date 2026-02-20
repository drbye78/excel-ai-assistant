import * as React from "react";
import { createRoot } from "react-dom/client";
import { ThemeProvider, initializeIcons } from "@fluentui/react";
import { App } from "./App";
import { AppProvider } from "@/context/AppContext";

// Initialize Fluent UI icons
initializeIcons();

// Office.js initialization
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const container = document.getElementById("root");
    if (container) {
      const root = createRoot(container);
      root.render(
        <ThemeProvider>
          <AppProvider>
            <App />
          </AppProvider>
        </ThemeProvider>
      );
    }
  }
});
