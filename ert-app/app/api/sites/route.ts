import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { guard, requireUser, siteScope } from "@/lib/auth";
export const dynamic = "force-dynamic";

export const GET = guard(async () => {
  const scope = siteScope(await requireUser());
  return NextResponse.json(await prisma.site.findMany({ where: scope ? { id: scope } : {}, orderBy: { name: "asc" }, include: { _count: { select: { employees: true } } } }));
});

export const POST = guard(async (req: NextRequest) => {
  await requireUser("SUPER_ADMIN");
  const b = await req.json();
  if (!b.name || !b.senderId) return NextResponse.json({ error: "Name and Sender ID are required." }, { status: 400 });
  try { return NextResponse.json(await prisma.site.create({ data: { name: String(b.name).trim(), senderId: String(b.senderId ?? "").trim() } }), { status: 201 }); }
  catch { return NextResponse.json({ error: "That name already exists." }, { status: 409 }); }
});