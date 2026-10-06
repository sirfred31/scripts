import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { clientIp, guard, requireAnyUser, setSessionCookie, signToken } from "@/lib/auth";
import { audit } from "@/lib/audit";
import { clear, hit, isLimited } from "@/lib/ratelimit";
import { decrypt, hashRecovery, newRecoveryCode, verifyTotp } from "@/lib/totp";

export const POST = guard(async (req: NextRequest) => {
  const s = await requireAnyUser();
  const k = `mfa-enable:${s.id}`;
  if (isLimited(k, 5)) return NextResponse.json({ error: "Too many tries. Wait 15 minutes." }, { status: 429 });
  const { code = "" } = await req.json().catch(() => ({}));
  const u = await prisma.user.findUniqueOrThrow({ where: { id: s.id } });
  if (u.mfaEnabled || !u.mfaSecret) return NextResponse.json({ error: "Start setup first." }, { status: 400 });

  const step = verifyTotp(decrypt(u.mfaSecret), String(code).replace(/\s/g, ""), -1);
  if (step === null) { hit(k, 15 * 60_000); return NextResponse.json({ error: "That code didn't match. Check the code in your app and try again." }, { status: 400 }); }
  clear(k);

  const codes = Array.from({ length: 8 }, newRecoveryCode);
  const upd = await prisma.user.update({
    where: { id: u.id },
    data: { mfaEnabled: true, mfaLastStep: step, mfaRecovery: await Promise.all(codes.map(hashRecovery)), sessionVersion: { increment: 1 } },
  });
  await audit("mfa.enabled", { userId: u.id, email: u.email, ip: clientIp(req) });
  const res = NextResponse.json({ recoveryCodes: codes }); // shown once, never retrievable again
  setSessionCookie(res, await signToken(u.id, upd.sessionVersion)); // other devices are signed out, this one stays in
  return res;
});
