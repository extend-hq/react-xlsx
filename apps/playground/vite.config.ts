import path from "path";
import tailwindcss from "@tailwindcss/vite";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tsconfigPaths from "vite-tsconfig-paths";

export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        main: path.resolve(__dirname, "index.html"),
        vanilla: path.resolve(__dirname, "vanilla.html"),
      },
    },
  },
  worker: {
    format: "es",
  },
  plugins: [react(), tsconfigPaths(), tailwindcss()],
  resolve: {
    alias: [
      { find: /^@extend-ai\/react-xlsx\/vanilla$/, replacement: path.resolve(__dirname, "../../packages/react-xlsx/dist/vanilla.js") },
      { find: /^@extend-ai\/react-xlsx\/internal$/, replacement: path.resolve(__dirname, "../../packages/react-xlsx/dist/internal.js") },
      { find: /^@extend-ai\/react-xlsx$/, replacement: path.resolve(__dirname, "../../packages/react-xlsx/dist/index.js") },
      { find: "@", replacement: path.resolve(__dirname, "./src") },
    ],
  },
});
