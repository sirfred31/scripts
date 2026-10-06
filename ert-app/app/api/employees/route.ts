import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { normalizePhone } from "@/lib/globe";
import { guard, requireUser, siteScope } from "@/lib/auth";
export const dynamic = "force-dynamic";

export const GET = guard(async (req: NextRequest) => {
  const s = await requireUser();
  const q = req.nextUrl.searchParams;
  const siteId = siteScope(s) ?? (q.get("siteId") || undefined); // scoped users can't widen the filter
  const search = q.get("q")?.trim();
  const employees = await prisma.employee.findMany({
    where: {
      ...(siteId ? { siteId } : {}),
      ...(q.get("deptId") ? { deptId: q.get("deptId")! } : {}),
      ...(search ? { OR: [
        { firstName: { contains: search, mode: "insensitive" } },
        { lastName: { contains: search, mode: "insensitive" } },
        { phoneNumber: { contains: search.replace(/\D/g, "") || "x" } },
      ] } : {}),
    },
    include: { site: true, department: true },
    orderBy: [{ lastName: "asc" }, { firstName: "asc" }],
    take: 500,
  });
  return NextResponse.json(employees);
});

export const POST = guard(async (req: NextRequest) => {
  const s = await requireUser("SUPER_ADMIN", "SITE_COMMANDER");
  const b = await req.json();
  const siteId = siteScope(s) ?? b.siteId; // commanders can only add to their own site
  const phoneNumber = normalizePhone(String(b.phoneNumber ?? ""));
  if (!b.firstName || !b.lastName || !siteId || !b.deptId || !phoneNumber)
    return NextResponse.json({ error: "First name, last name, site, department and a valid PH mobile number are required." }, { status: 400 });
  if (!(await prisma.site.findUnique({ where: { id: siteId } })) || !(await prisma.department.findUnique({ where: { id: b.deptId } })))
    return NextResponse.json({ error: "Choose an existing site and department." }, { status: 400 });
  if (String(b.firstName).length > 80 || String(b.lastName).length > 80) return NextResponse.json({ error: "Names are limited to 80 characters." }, { status: 400 });
  if (await prisma.employee.findUnique({ where: { phoneNumber } }))
    return NextResponse.json({ error: "That phone number already belongs to another employee." }, { status: 409 });
  const e = await prisma.employee.create({ data: { firstName: b.firstName.trim(), lastName: b.lastName.trim(), phoneNumber, siteId, deptId: b.deptId } });
  return NextResponse.json(e, { status: 201 });
});
