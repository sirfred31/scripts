import { NextRequest, NextResponse } from "next/server";
import type { Role } from "@prisma/client";
import { prisma } from "@/lib/db";
import { clientIp, guard, hashPassword, MIN_PASSWORD, requireUser } from "@/lib/auth";
import { audit } from "@/lib/audit";

const ROLES: Role[] = ["SUPER_ADMIN", "SITE_COMMANDER", "OPERATOR"];

export const GET = guard(async () => {
  await requireUser("SUPER_ADMIN");
  return NextResponse.json(await prisma.user.findMany({ select: { id: true, email: true, role: true, active: true, siteId: true, mfaEnabled: true, site: { select: { name: true } } }, orderBy: { email: "asc" } }));
});

export const POST = guard(async (req: NextRequest) => {
  const me = await requireUser("SUPER_ADMIN");
  const b = await req.json();
  const email = String(b.email ?? "").trim().toLowerCase();
  const password = String(b.password ?? "");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 200) return NextResponse.json({ error: "Enter a valid email address." }, { status: 400 });
  if (password.length < MIN_PASSWORD || password.length > 200 || password.toLowerCase() === email)
    return NextResponse.json({ error: `Use a password of at least ${MIN_PASSWORD} characters.` }, { status: 400 });
  if (!ROLES.includes(b.role)) return NextResponse.json({ error: "Choose a role." }, { status: 400 });
  if (b.role !== "SUPER_ADMIN" && !(await prisma.site.findUnique({ where: { id: String(b.siteId ?? "") } })))
    return NextResponse.json({ error: "Site Commanders and Operators need a site." }, { status: 400 });
  try {
    const u = await prisma.user.create({
      data: { email, passwordHash: await hashPassword(password), role: b.role, siteId: b.role === "SUPER_ADMIN" ? null : b.siteId },
      select: { id: true, email: true, role: true, siteId: true },
    });
    await audit("user.create", { userId: me.id, email: me.email, ip: clientIp(req), detail: { target: u.id, role: u.role } });
    return NextResponse.json(u, { status: 201 });
  } catch { return NextResponse.json({ error: "That email is already in use." }, { status: 409 }); }
});
