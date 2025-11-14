import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      "/merge": {
        target: "http://localhost:8010",
        changeOrigin: true,
      },
      "/pdf": {
        target: "http://localhost:8010",
        changeOrigin: true,
      },
      "/master": {
        target: "http://localhost:8010",
        changeOrigin: true,
      },
    },
  },
});
