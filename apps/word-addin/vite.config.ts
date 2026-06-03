import fs from "node:fs";
import path from "node:path";
import { defineConfig } from "vite";

const certDir = path.join(process.env.USERPROFILE ?? "", ".office-addin-dev-certs");

export default defineConfig(({ mode }) => {
  const isProd = mode === "production";

  // In production: relative paths, no HTTPS, no proxy, agent on same origin
  if (isProd) {
    return {
      base: "./",
      build: {
        outDir: "dist",
        emptyOutDir: true,
      },
      define: {
        __AGENT_BASE__: '""',
      },
    };
  }

  // Development mode: HTTPS dev server with proxy to local-agent
  const httpsOptions = {
    cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
    key: fs.readFileSync(path.join(certDir, "localhost.key")),
  };

  return {
    server: {
      host: "localhost",
      port: 3001,
      strictPort: true,
      https: httpsOptions,
      proxy: {
        "/api": {
          target: "http://127.0.0.1:8787",
          changeOrigin: true,
          rewrite: (p) => p.replace(/^\/api/, ""),
        },
      },
    },
    define: {
      __AGENT_BASE__: '"/api"',
    },
  };
});
