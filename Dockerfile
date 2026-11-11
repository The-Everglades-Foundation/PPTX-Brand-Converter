# Production image: Astro Node SSR + python-pptx venv.
# Cloudflare Workers / Workers Builds cannot run this process (no child Python).
FROM node:22-bookworm-slim

RUN apt-get update \
  && apt-get install -y --no-install-recommends python3 python3-venv python3-pip \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY . .

RUN python3 -m venv converter/.venv \
  && converter/.venv/bin/pip install --no-cache-dir -r converter/requirements.txt \
  && npm run build \
  && npm cache clean --force

ENV HOST=0.0.0.0
ENV PORT=8080
ENV NODE_ENV=production

EXPOSE 8080

CMD ["node", "dist/server/entry.mjs"]
