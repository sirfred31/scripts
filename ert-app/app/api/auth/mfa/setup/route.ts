import { NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { guard, requireAnyUser } from "@/lib/auth";
import { encrypt, newSecret, otpauthUri } from "@/lib/totp";

// Starts setup: creates a fresh secret (not active until a code is confirmed) and returns a QR code.
export const POST = guard(async () => {
  const s = await requireAnyUser();
  const u = await prisma.user.findUniqueOrThrow({ where: { id: s.id } });
  if (u.mfaEnabled) return NextResponse.json({ error: "Two-step sign-in is already on." }, { status: 400 });
  const secret = newSecret();
  await prisma.user.update({ where: { id: u.id }, data: { mfaSecret: encrypt(secret) } });
  const uri = otpauthUri(secret, u.email);
  const QR = await import("qrcode");
  return NextResponse.json({ secret, qr: await QR.toDataURL(uri, { margin: 1, width: 220 }) });
});
