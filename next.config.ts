import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "standalone",
  experimental: {
    serverActions: {
      bodySizeLimit: "100mb",
    },
  },
  // Do not bundle these packages — require() them from node_modules at runtime.
  // Prevents webpack from choking on WASM, native bindings, or complex CJS/ESM packages.
  serverExternalPackages: ["@anthropic-ai/sdk", "mupdf", "pdf-lib", "exceljs"],
};

export default nextConfig;
