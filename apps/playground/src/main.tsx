import * as React from "react";
import { Analytics } from "@vercel/analytics/react";
import wasmUrl from "@extend-ai/react-xlsx/duke_sheets_wasm_bg.wasm?url";
import { setWasmSource } from "@extend-ai/react-xlsx";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { PlaygroundCustomizerProvider } from "./components/playground-customizer";
import { ThemeProvider } from "./components/theme-provider";
import "./index.css";

setWasmSource(wasmUrl);

createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <ThemeProvider attribute="class" defaultTheme="system" enableSystem disableTransitionOnChange>
      <PlaygroundCustomizerProvider>
        <App />
      </PlaygroundCustomizerProvider>
      <Analytics />
    </ThemeProvider>
  </React.StrictMode>
);
