import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

const root = import.meta.dirname;

// Dev:  http://localhost:5173/          (redirect page: /redirect.html)
// Prod: https://mailai.uk/app/          (redirect page: /app/redirect.html), built into website/app
export default defineConfig(({ command }) => ({
  base: command === "build" ? "/app/" : "/",
  plugins: [react()],
  build: {
    outDir: resolve(root, "../website/app"),
    emptyOutDir: true,
    rollupOptions: {
      input: {
        main: resolve(root, "index.html"),
        redirect: resolve(root, "redirect.html"), // MSAL v5 redirect bridge
      },
    },
  },
  server: {
    port: 5173,
    strictPort: true, // the Entra registration allows exactly this port
    // Same-origin API calls in dev, like production (mailai.uk/v2/* -> Worker).
    // API_TARGET=http://127.0.0.1:8787 to use a local `wrangler dev` Worker instead.
    proxy: { "/v2": { target: process.env.API_TARGET || "https://mailai.uk", changeOrigin: true } },
  },
}));
