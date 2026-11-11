import type { APIRoute } from "astro";
import { checkPassword, cookieName, signPassword } from "../../lib/auth";

export const prerender = false;

export const POST: APIRoute = async ({ request, cookies }) => {
  const form = await request.formData();
  const password = String(form.get("password") ?? "");
  if (!checkPassword(password)) {
    return new Response(JSON.stringify({ error: "That password is not correct." }), {
      status: 401,
      headers: { "Content-Type": "application/json" },
    });
  }
  const token = signPassword();
  if (token) {
    cookies.set(cookieName(), token, {
      httpOnly: true,
      sameSite: "lax",
      secure: process.env.NODE_ENV === "production",
      path: "/",
      maxAge: 60 * 60 * 12,
    });
  }
  return new Response(JSON.stringify({ ok: true }), {
    headers: { "Content-Type": "application/json" },
  });
};
