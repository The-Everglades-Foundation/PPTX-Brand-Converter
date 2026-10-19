import type { APIRoute } from "astro";
import { existsSync } from "node:fs";
import { pythonExecutable, TEMPLATE_PATH } from "../lib/run-converter";

export const prerender = false;

export const GET: APIRoute = async () => {
  let python = false;
  try {
    python = existsSync(pythonExecutable());
  } catch {
    python = false;
  }
  const template = existsSync(TEMPLATE_PATH);
  const ok = python && template;
  return new Response(JSON.stringify({ ok, python, template }), {
    status: ok ? 200 : 503,
    headers: { "Content-Type": "application/json" },
  });
};
