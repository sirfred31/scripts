import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { normalizePhone } from "@/lib/globe";
import { clientIp, guard, HttpError, requireUser, siteScope } from "@/lib/auth";
import { audit } from "@/lib/audit";

type Ctx = { params: Promise<{ id: string }> };

async function ownedEmployee(id: string) {
  const s = await requireUser("SUPER_ADMIN", "SITE_COMMANDER");
  const e = await prisma.employee.findUnique({ where: { id } });
  if (!e) throw new HttpError(404, "Employee not found.");
  const scope = siteScope(s);
  if (scope && e.siteId !== scope) throw new HttpError(403, "That employee belongs to another site.");
  return { s, scope };
}

export const PATCH = guard(async (req: NextRequest, { params }: Ctx) => {
  const { id } = await params;
  const { scope } = await ownedEmployee(id);
  const b = await req.json();
  const data: Record<string, string> = {};
  for (const k of ["firstName", "lastName", "deptId"]) if (b[k]) data[k] = String(b[k]).trim();
  if (b.siteId) data.siteId = scope ?? String(b.siteId); // commanders can't move people to other sites
  if (b.phoneNumber) {
    const p = normalizePhone(String(b.phoneNumber));
    if (!p) return NextResponse.json({ error: "Invalid PH mobile number." }, { status: 400 });
    data.phoneNumber = p;
  }
  try { return NextResponse.json(await prisma.employee.update({ where: { id }, data })); }
  catch { return NextResponse.json({ error: "Could not update. Check the phone number isn't already in use." }, { status: 409 }); }
});

export const DELETE = guard(async (_: NextRequest, { params }: Ctx) => {
  const { id } = await params;
  const { s } = await ownedEmployee(id);
  await audit("employee.delete", { userId: s.id, email: s.email, ip: clientIp(_), detail: { employeeId: id } });
  await prisma.$transaction([prisma.sMSResponse.deleteMany({ where: { employeeId: id } }), prisma.employee.delete({ where: { id } })]);
  return NextResponse.json({ ok: true });
});
