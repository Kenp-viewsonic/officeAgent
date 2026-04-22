import fs from "node:fs";
import path from "node:path";
import { defineConfig } from "vite";

const certDir = path.join(process.env.USERPROFILE ?? "", ".office-addin-dev-certs");
const httpsOptions = {
  cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
  key: fs.readFileSync(path.join(certDir, "localhost.key")),
};

export default defineConfig({
  server: {
    host: "localhost",
    port: 3001,
    strictPort: true,
    https: httpsOptions,
    proxy: {
      "/api": {
        target: "http://127.0.0.1:8787",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api/, ""),
      },
    },
  },
});
