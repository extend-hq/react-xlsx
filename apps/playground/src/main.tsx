import * as React from "react";
import { Analytics } from "@vercel/analytics/react";
import { setWasmSource } from "@extend-ai/react-xlsx";
import { ThemeProvider } from "next-themes";
import { createRoot } from "react-dom/client";
import { TooltipProvider } from "@/components/ui/tooltip";
import { App } from "./App";
import "./index.css";

setWasmSource("/duke_sheets_wasm_bg.wasm");

createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <ThemeProvider attribute="class" defaultTheme="light" disableTransitionOnChange enableSystem={false}>
      <TooltipProvider delay={400}>
        <App />
      </TooltipProvider>
      <Analytics />
    </ThemeProvider>
  </React.StrictMode>
);
