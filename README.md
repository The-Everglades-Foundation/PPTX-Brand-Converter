# PPTX-Brand-Converter

Internal Everglades Foundation staff tool: upload any PowerPoint, inspect its content, rebuild every slide onto the Foundation template, and download a branded deck plus a QA report.

## Run locally

Requires **Node.js 22.12** or newer and **Python 3.11+**. For visual QA, install **LibreOffice** and **Poppler** (`pdftoppm`).

```sh
python -m venv converter/.venv
# macOS / Linux
converter/.venv/bin/python -m pip install -r converter/requirements.txt
# Windows
converter/.venv\Scripts\python -m pip install -r converter/requirements.txt
npm install
npm run dev
```

Dev server: `http://127.0.0.1:4321`

Optional staff password: copy `.env.example` to `.env` and set `INTERNAL_PASSWORD`.

Place Communications’ official 2023 template at `templates/everglades.pptx`. If you only have a `.potx`, the engine rewrites it to `.pptx` on inspect. That file contains **both** palette colors — **Green** is Sawgrass Lime `C1D451`, **Blue** is Water Teal `00ACBF`. Staff pick one; convert applies that colorway to the whole deck.

## Commands

| Command            | Action |
| :----------------- | :----- |
| `npm install`      | Install dependencies |
| `npm run dev`      | Dev server at `localhost:4321` |
| `npm run build`    | Production build to `./dist/` |
| `npm run start`    | Run the production Node server (`HOST` / `PORT`) |
| `npm run preview`  | Preview the production build locally |
| `npm run template` | Rebuild a python-pptx starter deck (dev only; does not replace the official Communications file) |
| `npm run fixture`  | Rebuild `fixtures/sample.pptx` |
| `rebrand/run.sh`   | Inspect → plan → build → QA (`COLORWAY=green|blue`) |

CLI (from repo root, using the converter venv):

```sh
python rebrand/inspect_template.py templates/everglades.pptx --out-dir .tmp/rebrand --colorway green
python rebrand/inspect_source.py fixtures/sample.pptx --out-dir .tmp/rebrand
python rebrand/plan.py --out-dir .tmp/rebrand --colorway green
python rebrand/build.py --out-dir .tmp/rebrand --colorway green
python rebrand/qa.py --out-dir .tmp/rebrand --colorway green
```

Or: `COLORWAY=blue rebrand/run.sh templates/everglades.pptx fixtures/sample.pptx .tmp/rebrand`

## Project structure

```text
rebrand/                 inspect → plan → rebuild → QA engine
converter/               template/fixture generators (not the conversion core)
templates/everglades.pptx
src/pages/index.astro    Staff upload UI
src/pages/api/           parse + convert (plan.json + QA flags)
src/layouts/Layout.astro
src/styles/global.css    Everglades color tokens
```

## Production (Node + Python)

Convert **spawns** `converter/.venv` and `rebrand/api.py` (`python-pptx`). That is a long-running Node process plus a Python interpreter. It will not run on **Cloudflare Workers** or **Workers Builds** (the `pptx-brand-converter` Worker check that fails with a missing entry point is the wrong host — ignore it).

Smallest reliable ship path is **Docker**:

```sh
docker build -t pptx-brand-converter .
docker run --rm -p 8080:8080 -e INTERNAL_PASSWORD pptx-brand-converter
```

Or `docker compose up --build` if the Compose plugin is installed. Then open `http://127.0.0.1:8080`. `GET /health` should return `{"ok":true,"python":true,"template":true}`.

Set `INTERNAL_PASSWORD` in the environment (or a `.env` next to `docker-compose.yml`) before exposing the service.

### Host it on Render

1. Open [Render Blueprints](https://dashboard.render.com/blueprints) and sign in (or create a Render account).
2. **New Blueprint** → connect `evergladesfoundation/PPTX-Brand-Converter`.
3. Apply `render.yaml`. Use at least the **Standard** plan (512 MB is tight for 50 MB decks).
4. In the service **Environment**, set `INTERNAL_PASSWORD`.
5. Wait for the deploy, then open the `*.onrender.com` URL and sign in.

Any other Docker host (Fly.io, a VM, Kubernetes) works the same: build this `Dockerfile`, publish port `8080`, set `HOST=0.0.0.0` and `INTERNAL_PASSWORD`.

Bare metal without Docker:

```sh
python3 -m venv converter/.venv
converter/.venv/bin/python -m pip install -r converter/requirements.txt
npm ci
npm run build
HOST=0.0.0.0 PORT=8080 NODE_ENV=production npm start
```

## Tailwind notes

This project uses the official Vite plugin (`@tailwindcss/vite`). Import `src/styles/global.css` from the layout — never from every page.
