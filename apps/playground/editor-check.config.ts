import { defineConfig } from "vite";
import config from "./vite.config";
export default defineConfig({ ...config, cacheDir: "../../output/playwright/cell-editor/vite-cache", server: { hmr: false } });
