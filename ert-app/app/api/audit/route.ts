import { NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { guard, requireUser } from "@/lib/auth";
export const dynamic = "force-dynamic";

export const GET = guard(async () => {
  await requireUser("SUPER_ADMIN");
  return NextResponse.json(await prisma.auditLog.findMany({ orderBy: { at: "desc" }, take: 100 }));
});
