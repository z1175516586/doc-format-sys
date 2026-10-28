import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

import { DEV_SERVER_PORT, API_PROXY_TARGET } from "./constants.ts";

export default defineConfig({
  plugins: [react()],
  server: {
    host: "127.0.0.1",
    port: DEV_SERVER_PORT,
    proxy: {
      "/api": {
        target: API_PROXY_TARGET,
        changeOrigin: true,
      },
    },
  },
});
