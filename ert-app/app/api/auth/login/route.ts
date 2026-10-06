import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { checkPassword, clientIp, dummyHash, setSessionCookie, signTicket, signToken } from "@/lib/auth";
import { clear, hit, isLimited } from "@/lib/ratelimit";
import { audit } from "@/lib/audit";

const WINDOW = 15 * 60_000;

export async function POST(req: NextRequest) {
  const ip = clientIp(req);
  const body = await req.json().catch(() => ({}));
  const email = String(body.email ?? "").trim().toLowerCase().slice(0, 200);
  const password = String(body.password ?? "").slice(0, 200);
  // Keyed by IP + email so a stranger can't lock a real user out from elsewhere; a separate per-IP cap stops spraying.
  const pair = `login:${ip}:${email}`, perIp = `login-ip:${ip}`;
  if (isLimited(pair, 5) || isLimited(perIp, 30))
    return NextResponse.json({ error: "Too many attempts. Try again in 15 minutes." }, { status: 429 });

  const u = await prisma.user.findUnique({ where: { email } });
  const ok = await checkPassword(password, u?.passwordHash ?? (await dummyHash())); // always runs bcrypt
  if (!u || !u.active || !ok) {
    hit(pair, WINDOW); hit(perIp, WINDOW);
    await audit("login.fail", { email, ip });
    return NextResponse.json({ error: "Email or password is incorrect." }, { status: 401 });
  }
  clear(pair);
  if (u.mfaEnabled) {
    // Password was right, but no session yet: the second step proves the authenticator code.
    await audit("login.password_ok", { userId: u.id, email, ip });
    return NextResponse.json({ mfaRequired: true, ticket: await signTicket(u.id) });
  }
  await audit("login.ok", { userId: u.id, email, ip });
  const res = NextResponse.json({ ok: true });
  setSessionCookie(res, await signToken(u.id, u.sessionVersion));
  return res;
}
