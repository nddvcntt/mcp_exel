# ── Stage 1: Builder ─────────────────────────────────────
FROM node:20-bookworm-slim AS builder

WORKDIR /build

COPY package*.json ./
RUN npm ci --omit=dev

# ── Stage 2: Python/uvx layer ────────────────────────────
FROM python:3.12-slim AS python-base

RUN pip install --no-cache-dir uv

# ── Stage 3: Final runner ────────────────────────────────
FROM node:20-bookworm-slim

WORKDIR /app

# Cài Python + uv để chạy uvx excel-mcp-server
RUN apt-get update \
  && apt-get install -y --no-install-recommends python3 python3-pip ca-certificates \
  && python3 -m pip install --break-system-packages --no-cache-dir uv \
  && rm -rf /var/lib/apt/lists/*

# Copy node_modules từ builder (không build lại)
COPY --from=builder /build/node_modules ./node_modules

COPY package*.json ./
COPY src ./src

ENV NODE_ENV=production \
    PORT=5003 \
    EXCEL_FILES_PATH=/data/excel_files \
    OUTPUT_DIR=/data/output \
    EXCEL_MCP_INTERNAL_PORT=5013 \
    EXCEL_MCP_SHELL=false \
    FILE_TTL_DAYS=5

RUN mkdir -p /data/excel_files /data/output

EXPOSE 5003

HEALTHCHECK --interval=30s --timeout=5s --start-period=60s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:5003/health').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"

CMD ["node", "src/index.js"]
