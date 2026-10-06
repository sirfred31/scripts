import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { guard, requireUser, siteScope } from "@/lib/auth";
export const dynamic = "force-dynamic";

export const GET = guard(async (req: NextRequest) => {
  const scope = siteScope(await requireUser());
  const q = req.nextUrl.searchParams;
  const siteId = scope ?? (q.get("siteId") || undefined);
  const visible = scope ? { OR: [{ siteId: null }, { siteId: scope }] } : {}; // never another site's campaign
  const campaign = q.get("campaignId")
    ? await prisma.broadcastCampaign.findFirst({ where: { id: q.get("campaignId")!, ...visible } })
    : await prisma.broadcastCampaign.findFirst({ where: visible, orderBy: { createdAt: "desc" } });
  const sites = await prisma.site.findMany({ where: scope ? { id: scope } : {}, orderBy: { name: "asc" } });
  if (!campaign) return NextResponse.json({ campaign: null, sites, counts: null, feed: [], unresponsive: [] });

  const where = { campaignId: campaign.id, ...(siteId ? { employee: { siteId } } : {}) };
  const grouped = await prisma.sMSResponse.groupBy({ by: ["status"], where, _count: true });
  const n = (s: string) => grouped.find((g) => g.status === s)?._count ?? 0;
  const counts = { total: grouped.reduce((a, g) => a + g._count, 0), safe: n("SAFE"), notSafe: n("NOT_SAFE"), help: n("HELP"), pending: n("PENDING") };

  const include = { employee: { include: { site: true, department: true } } };
  const [feed, unresponsive] = await Promise.all([
    prisma.sMSResponse.findMany({ where: { ...where, status: { not: "PENDING" as const } }, include, orderBy: { updatedAt: "desc" }, take: 50 }),
    prisma.sMSResponse.findMany({ where: { ...where, status: "PENDING" as const }, include, take: 200 }),
  ]);
  return NextResponse.json({ campaign, sites, counts, feed, unresponsive });
});