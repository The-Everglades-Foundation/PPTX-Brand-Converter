// @ts-check
import node from "@astrojs/node";
import tailwindcss from "@tailwindcss/vite";
import { defineConfig } from "astro/config";

export default defineConfig({
  site: "https://github.com/evergladesfoundation/PPTX-Brand-Converter",
  output: "server",
  adapter: node({ mode: "standalone" }),
  // Internal staff app. Astro’s CSRF origin check 403s multipart POSTs from
  // Cursor Simple Browser / Cloud Agent port-forwards (Origin ≠ 127.0.0.1),
  // which left the UI stuck on “Reading slides…”. Auth is the staff password.
  security: {
    checkOrigin: false,
  },
  server: {
    host: "127.0.0.1",
  },
  vite: {
    plugins: [tailwindcss()],
  },
});
