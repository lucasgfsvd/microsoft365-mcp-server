# syntax=docker/dockerfile:1.7

# ---- Build stage ----
FROM node:25-alpine AS build
WORKDIR /app

# Install dependencies (including dev deps for tsup/typescript build)
COPY package.json package-lock.json* ./
RUN --mount=type=cache,target=/root/.npm \
    npm install --no-audit --no-fund

# Copy sources and build
COPY tsconfig.json tsup.config.ts ./
COPY src ./src
RUN npm run build

# Prune to production deps for final image
RUN npm prune --omit=dev

# ---- Runtime stage ----
FROM gcr.io/distroless/nodejs20-debian12:nonroot AS runtime
WORKDIR /app

ENV NODE_ENV=production \
    MCP_TOKEN_CACHE_PATH=/data/tokencache.json

COPY --from=build /app/node_modules ./node_modules
COPY --from=build /app/dist ./dist
COPY --from=build /app/package.json ./package.json

# Token cache volume
VOLUME ["/data"]

USER nonroot
ENTRYPOINT ["/nodejs/bin/node", "/app/dist/index.js"]
