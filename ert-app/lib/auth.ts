import { SignJWT, jwtVerify } from "jose";
import bcrypt from "bcryptjs";
import { cookies } from "next/headers";
import { NextResponse } from "next/server";
import type { Role } from "@prisma/client";
import { prisma } from "./db";

export const COOKIE = "ert_session";
export const MAX_AGE = 60 * 60 * 12; // 12 hours

// Fail closed: a missing or short secret must never produce a usable (forgeable) key.
const key = () => {
  const s = process.env.AUTH_SECRET;
  if (!s || s.length < 32) throw new Error("AUTH_SECRET must be set to at least 32 characters.");
  return new TextEncoder().encode(s);
};

export const MIN_PASSWORD = 12;
export const hashPassword = (p: string) => bcrypt.hash(p, 12);
export const checkPassword = (p: string, h: string) => bcrypt.compare(p, h);
let dummy: string | undefined;
/** Compared against when the email is unknown, so response time doesn't reveal which emails exist. */
export const dummyHash = async () => (dummy ??= await hashPassword("not-a-real-password"));

export const signToken = (uid: string, version: number) =>
  new SignJWT({ v: version }).setProtectedHeader({ alg: "HS256" }).setSubject(uid).setExpirationTime(`${MAX_AGE}s`).sign(key());

/** Short-lived ticket proving the password step passed; only good for the two-step code check. */
export const signTicket = (uid: string) =>
  new SignJWT({ p: "mfa" }).setProtectedHeader({ alg: "HS256" }).setSubject(uid).setExpirationTime("5m").sign(key());
export async function readTicket(token: string) {
  try { const { payload } = await jwtVerify(token, key(), { algorithms: ["HS256"] }); return payload.p === "mfa" ? payload.sub ?? null : null; }
  catch { return null; }
}

export function setSessionCookie(res: NextResponse, token: string) {
  res.cookies.set(COOKIE, token, { httpOnly: true, sameSite: "lax", secure: process.env.NODE_ENV === "production", path: "/", maxAge: MAX_AGE });
}

export type Session = { id: string; email: string; role: Role; siteId: string | null; mfaEnabled: boolean };

/** Re-checks the user in the DB: deactivation, role changes and password resets take effect on the next request. */
export async function getSession(): Promise<Session | null> {
  const token = (await cookies()).get(COOKIE)?.value;
  if (!token) return null;
  try {
    const { payload } = await jwtVerify(token, key(), { algorithms: ["HS256"] });
    const u = await prisma.user.findUnique({ where: { id: payload.sub! } });
    if (payload.p || !u?.active || payload.v !== u.sessionVersion) return null; // tickets are never sessions
    return { id: u.id, email: u.email, role: u.role, siteId: u.siteId, mfaEnabled: u.mfaEnabled };
  } catch { return null; }
}

export class HttpError extends Error { constructor(public status: number, msg: string, public code?: string) { super(msg); } }

/** Signed in, but without the Super Admin two-step requirement. Only for the MFA setup endpoints. */
export async function requireAnyUser() {
  const s = await getSession();
  if (!s) throw new HttpError(401, "Sign in required.");
  return s;
}

export async function requireUser(...roles: Role[]) {
  const s = await requireAnyUser();
  if (s.role === "SUPER_ADMIN" && !s.mfaEnabled) throw new HttpError(403, "Set up two-step sign-in to continue.", "MFA_SETUP_REQUIRED");
  if (roles.length && !roles.includes(s.role)) throw new HttpError(403, "You don't have permission to do that.");
  return s;
}

/** undefined = all sites (Super Admin). Otherwise the one site this user may touch ("none" matches nothing). */
export const siteScope = (s: Session) => (s.role === "SUPER_ADMIN" ? undefined : s.siteId ?? "none");

/** Only trustworthy behind your own reverse proxy that sets x-forwarded-for. */
export const clientIp = (req: Request) => req.headers.get("x-forwarded-for")?.split(",")[0].trim() || "unknown";

export const guard = <A extends unknown[]>(fn: (...a: A) => Promise<Response>) => async (...a: A) => {
  try { return await fn(...a); }
  catch (e) {
    if (e instanceof HttpError) return NextResponse.json({ error: e.message, code: e.code }, { status: e.status });
    console.error(e); // details stay in server logs, never in the response
    return NextResponse.json({ error: "Something went wrong. Try again." }, { status: 500 });
  }
};
