import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { clientIp, readTicket, setSessionCookie, signToken } from "@/lib/auth";
import { audit } from "@/lib/audit";
import { clear, hit, isLimited } from "@/lib/ratelimit";
import { decrypt, matchRecovery, verifyTotp } from "@/lib/totp";

const WINDOW = 15 * 60_000;

// Second sign-in step. Accepts an authenticator code or a one-time recovery code.
export async function POST(req: NextRequest) {
  const ip = clientIp(req);
  const { ticket = "", code = "" } = await req.json().catch(() => ({}));
  const uid = await readTicket(String(ticket));
  if (!uid) return NextResponse.json({ error: "That sign-in expired. Start again." }, { status: 401 });

  const k = `mfa:${uid}`, kip = `mfa-ip:${ip}`;
  if (isLimited(k, 5) || isLimited(kip, 30)) return NextResponse.json({ error: "Too many attempts. Try again in 15 minutes." }, { status: 429 });

  const u = await prisma.user.findUnique({ where: { id: uid } });
  if (!u || !u.active || !u.mfaEnabled || !u.mfaSecret) return NextResponse.json({ error: "That sign-in expired. Start again." }, { status: 401 });

  const c = String(code).trim().slice(0, 32);
  let ok = false, usedRecovery = false;
  const digits = c.replace(/\s/g, "");
  if (/^\d{6}$/.test(digits)) {
    const step = verifyTotp(decrypt(u.mfaSecret), digits, u.mfaLastStep ?? -1);
    // Atomic: only the first request to claim a step wins, so a code can't be replayed or raced.
    if (step !== null) ok = (await prisma.user.updateMany({ where: { id: u.id, OR: [{ mfaLastStep: null }, { mfaLastStep: { lt: step } }] }, data: { mfaLastStep: step } })).count === 1;
  } else {
    const i = await matchRecovery(c, u.mfaRecovery);
    if (i >= 0) {
      const left = u.mfaRecovery.filter((_, j) => j !== i);
      ok = (await prisma.user.updateMany({ where: { id: u.id, mfaRecovery: { equals: u.mfaRecovery } }, data: { mfaRecovery: left } })).count === 1;
      usedRecovery = ok;
    }
  }
  if (!ok) {
    hit(k, WINDOW); hit(kip, WINDOW);
    await audit("mfa.fail", { userId: u.id, email: u.email, ip });
    return NextResponse.json({ error: "That code didn't work. Try the newest code, or a recovery code." }, { status: 401 });
  }
  clear(k);
  await audit(usedRecovery ? "login.ok.recovery" : "login.ok", { userId: u.id, email: u.email, ip, detail: { mfa: true, recoveryLeft: usedRecovery ? u.mfaRecovery.length - 1 : undefined } });
  const res = NextResponse.json({ ok: true });
  setSessionCookie(res, await signToken(u.id, u.sessionVersion));
  return res;
}
