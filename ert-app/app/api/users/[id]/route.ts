import { NextRequest, NextResponse } from "next/server";
import type { Role } from "@prisma/client";
import { prisma } from "@/lib/db";
import { clientIp, guard, hashPassword, MIN_PASSWORD, requireUser } from "@/lib/auth";
import { audit } from "@/lib/audit";

type Ctx = { params: Promise<{ id: string }> };
const ROLES: Role[] = ["SUPER_ADMIN", "SITE_COMMANDER", "OPERATOR"];

// Body may contain: active, role, siteId, password (reset), resetMfa. Any change signs that user out everywhere.
export const PATCH = guard(async (req: NextRequest, { params }: Ctx) => {
  const me = await requireUser("SUPER_ADMIN");
  const { id } = await params;
  const b = await req.json();
  const target = await prisma.user.findUnique({ where: { id } });
  if (!target) return NextResponse.json({ error: "User not found." }, { status: 404 });

  const losesAdmin = target.role === "SUPER_ADMIN" && (b.active === false || (b.role && b.role !== "SUPER_ADMIN"));
  if (losesAdmin) {
    if (id === me.id) return NextResponse.json({ error: "You can't deactivate or demote your own account." }, { status: 400 });
    if (!(await prisma.user.count({ where: { role: "SUPER_ADMIN", active: true, id: { not: id } } })))
      return NextResponse.json({ error: "At least one active Super Admin is required." }, { status: 400 });
  }

  const data: Record<string, unknown> = {};
  if (typeof b.active === "boolean") data.active = b.active;
  if (b.role) { if (!ROLES.includes(b.role)) return NextResponse.json({ error: "Choose a valid role." }, { status: 400 }); data.role = b.role; }
  if ("siteId" in b) {
    if (b.siteId && !(await prisma.site.findUnique({ where: { id: String(b.siteId) } }))) return NextResponse.json({ error: "Site not found." }, { status: 400 });
    data.siteId = b.siteId || null;
  }
  const finalRole = (data.role as Role) ?? target.role;
  const finalSite = "siteId" in data ? data.siteId : target.siteId;
  if (finalRole !== "SUPER_ADMIN" && !finalSite) return NextResponse.json({ error: "Site Commanders and Operators need a site." }, { status: 400 });
  if (b.password) {
    const pw = String(b.password);
    if (pw.length < MIN_PASSWORD || pw.length > 200) return NextResponse.json({ error: `Use a password of at least ${MIN_PASSWORD} characters.` }, { status: 400 });
    data.passwordHash = await hashPassword(pw);
  }
  if (b.resetMfa) { // lost phone: the user sets up a new authenticator at next sign-in
    if (id === me.id) return NextResponse.json({ error: "Ask another Super Admin to reset your two-step sign-in." }, { status: 400 });
    Object.assign(data, { mfaEnabled: false, mfaSecret: null, mfaRecovery: [], mfaLastStep: null });
  }
  data.sessionVersion = { increment: 1 };
  await prisma.user.update({ where: { id }, data });
  await audit("user.update", { userId: me.id, email: me.email, ip: clientIp(req), detail: { target: id, fields: Object.keys(data).filter((k) => k !== "passwordHash" && k !== "sessionVersion"), passwordReset: !!b.password, mfaReset: !!b.resetMfa } });
  return NextResponse.json({ ok: true });
});
