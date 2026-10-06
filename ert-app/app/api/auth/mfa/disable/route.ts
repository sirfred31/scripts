import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { clientIp, guard, requireAnyUser, setSessionCookie, signToken } from "@/lib/auth";
import { audit } from "@/lib/audit";
import { hit, isLimited } from "@/lib/ratelimit";
import { decrypt, verifyTotp } from "@/lib/totp";

// Operators and Site Commanders may turn two-step off (with a current code). Super Admins can't.
export const POST = guard(async (req: NextRequest) => {
  const s = await requireAnyUser();
  if (s.role === "SUPER_ADMIN") return NextResponse.json({ error: "Super Admins must keep two-step sign-in on." }, { status: 403 });
  const k = `mfa-disable:${s.id}`;
  if (isLimited(k, 5)) return NextResponse.json({ error: "Too many tries. Wait 15 minutes." }, { status: 429 });
  const { code = "" } = await req.json().catch(() => ({}));
  const u = await prisma.user.findUniqueOrThrow({ where: { id: s.id } });
  if (!u.mfaEnabled || !u.mfaSecret) return NextResponse.json({ error: "Two-step sign-in is already off." }, { status: 400 });
  if (verifyTotp(decrypt(u.mfaSecret), String(code).replace(/\s/g, ""), u.mfaLastStep ?? -1) === null) {
    hit(k, 15 * 60_000);
    return NextResponse.json({ error: "That code didn't match." }, { status: 400 });
  }
  const upd = await prisma.user.update({ where: { id: u.id }, data: { mfaEnabled: false, mfaSecret: null, mfaRecovery: [], mfaLastStep: null, sessionVersion: { increment: 1 } } });
  await audit("mfa.disabled", { userId: u.id, email: u.email, ip: clientIp(req) });
  const res = NextResponse.json({ ok: true });
  setSessionCookie(res, await signToken(u.id, upd.sessionVersion));
  return res;
});
