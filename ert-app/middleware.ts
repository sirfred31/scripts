import { NextRequest, NextResponse } from "next/server";
import { jwtVerify } from "jose";

// Gate only: signature check. Routes re-check the user and role via lib/auth.
const PUBLIC = ["/login", "/api/auth/login", "/api/auth/mfa/verify", "/api/webhooks/"]; // the webhook has its own secret
const WRITE = new Set(["POST", "PUT", "PATCH", "DELETE"]);

export async function middleware(req: NextRequest) {
  const p = req.nextUrl.pathname;
  const isApi = p.startsWith("/api/");

  // CSRF defence in depth (cookies are also SameSite=Lax): reject cross-site writes to the API.
  if (isApi && WRITE.has(req.method) && !p.startsWith("/api/webhooks/")) {
    const origin = req.headers.get("origin");
    if ((origin && new URL(origin).host !== req.headers.get("host")) || req.headers.get("sec-fetch-site") === "cross-site")
      return NextResponse.json({ error: "Cross-site request blocked." }, { status: 403 });
  }

  if (PUBLIC.some((x) => p.startsWith(x))) return NextResponse.next();
  const t = req.cookies.get("ert_session")?.value;
  const secret = process.env.AUTH_SECRET;
  if (t && secret && secret.length >= 32) {
    try { await jwtVerify(t, new TextEncoder().encode(secret), { algorithms: ["HS256"] }); return NextResponse.next(); } catch {}
  }
  return isApi ? NextResponse.json({ error: "Sign in required." }, { status: 401 }) : NextResponse.redirect(new URL("/login", req.url));
}
export const config = { matcher: ["/((?!_next/|favicon.ico).*)"] };
