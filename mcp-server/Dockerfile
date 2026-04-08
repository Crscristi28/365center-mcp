FROM node:22-alpine AS builder

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY tsconfig.json ./
COPY src/ ./src/

RUN npx tsc

# --- Production image ---
FROM node:22-alpine

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY --from=builder /app/dist/ ./dist/

# Run as non-root user for security
RUN addgroup -S mcp && adduser -S mcp -G mcp -h /home/mcp
RUN mkdir -p /home/mcp/.365center-mcp && chown -R mcp:mcp /home/mcp
USER mcp
ENV HOME=/home/mcp

# MCP servers communicate via stdio
ENTRYPOINT ["node", "dist/index.js"]
