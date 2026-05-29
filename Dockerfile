FROM node:20-alpine AS base
WORKDIR /app

# ── Install dependencies ──────────────────────────────────────────────────────
FROM base AS deps
COPY package.json package-lock.json ./
RUN npm ci

# ── Build ─────────────────────────────────────────────────────────────────────
FROM base AS builder
ENV NEXT_TELEMETRY_DISABLED=1
COPY --from=deps /app/node_modules ./node_modules
COPY . .
RUN npm run build

# ── Runtime image ─────────────────────────────────────────────────────────────
FROM base AS runner
ENV NODE_ENV=production
ENV NEXT_TELEMETRY_DISABLED=1
# Railway injects PORT; default to 3000 for local Docker runs
ENV PORT=3000
ENV HOSTNAME="0.0.0.0"

# Standalone output — minimal self-contained server
COPY --from=builder /app/public ./public
COPY --from=builder /app/.next/standalone ./
COPY --from=builder /app/.next/static ./.next/static

# mupdf ships .wasm files that @vercel/nft may not trace — copy the full package
COPY --from=deps /app/node_modules/mupdf ./node_modules/mupdf

EXPOSE 3000
CMD ["node", "server.js"]
