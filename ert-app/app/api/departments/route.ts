import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { guard, requireUser, siteScope } from "@/lib/auth";
export const dynamic = "force-dynamic";

export const GET = guard(async () => {
  await requireUser();
  return NextResponse.json(await prisma.department.findMany({ orderBy: { name: "asc" }, include: { _count: { select: { employees: true } } } }));
});

export const POST = guard(async (req: NextRequest) => {
  await requireUser("SUPER_ADMIN");
  const b = await req.json();
  if (!b.name ) return NextResponse.json({ error: "Name is required." }, { status: 400 });
  try { return NextResponse.json(await prisma.department.create({ data: { name: String(b.name).trim() } }), { status: 201 }); }
  catch { return NextResponse.json({ error: "That name already exists." }, { status: 409 }); }
});