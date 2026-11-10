import { defineConfig } from "vite";
import config from "./vite.config";
export default defineConfig({ ...config, server: { hmr: false } });
