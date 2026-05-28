import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  experimental: {
    serverActions: {
      bodySizeLimit: "20mb",
    },
  },
  webpack: (config, { isServer }) => {
    if (isServer) {
      // canvas is a native module — must not be bundled, require()d at runtime
      // pdfjs-dist is handled via /* webpackIgnore: true */ at the import site
      if (Array.isArray(config.externals)) {
        config.externals = [...config.externals, "canvas"];
      } else if (typeof config.externals === "function") {
        config.externals = [config.externals, "canvas"];
      } else if (config.externals) {
        config.externals = [config.externals, "canvas"];
      } else {
        config.externals = ["canvas"];
      }
    }
    return config;
  },
};

export default nextConfig;
