import type { NextConfig } from "next";
import path from "path";

const nextConfig: NextConfig = {
  experimental: {
    serverActions: {
      bodySizeLimit: "20mb",
    },
  },
  webpack: (config, { isServer }) => {
    // Use legacy build of pdfjs-dist in Node.js (avoids DOMMatrix error)
    config.resolve = {
      ...config.resolve,
      alias: {
        ...(config.resolve?.alias ?? {}),
        "pdfjs-dist": path.join(
          process.cwd(),
          "node_modules/pdfjs-dist/legacy/build/pdf.mjs"
        ),
      },
    };

    if (isServer) {
      // canvas is a native module — must be required at runtime, not bundled
      const ext = Array.isArray(config.externals)
        ? config.externals
        : config.externals
        ? [config.externals]
        : [];
      config.externals = [...ext, "canvas"];
    }

    return config;
  },
};

export default nextConfig;
