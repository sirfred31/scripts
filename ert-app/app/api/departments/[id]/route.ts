import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { guard, requireUser } from "@/lib/auth";

export const PATCH = guard(async (req: NextRequest, { params }: { params: Promise<{ id: string }> }) => {
  await requireUser("SUPER_ADMIN");
  const { id } = await params;
  const b = await req.json();
  const data: Record<string, string> = {};
  if (b.name) data.name = String(b.name).trim();

  try { return NextResponse.json(await prisma.department.update({ where: { id }, data })); }
  catch { return NextResponse.json({ error: "Could not update. The name may already exist." }, { status: 409 }); }
});

export const DELETE = guard(async (_: NextRequest, { params }: { params: Promise<{ id: string }> }) => {
  await requireUser("SUPER_ADMIN");
  const { id } = await params;
  const n = await prisma.employee.count({ where: { deptId: id } });
  if (n) return NextResponse.json({ error: `Move or remove its ${n} employees first.` }, { status: 409 });
  await prisma.department.delete({ where: { id } });
  return NextResponse.json({ ok: true });
});
