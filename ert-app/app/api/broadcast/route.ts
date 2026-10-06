import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { publish } from "@/lib/events";
import { sendSms } from "@/lib/globe";
import { clientIp, guard, requireUser, siteScope } from "@/lib/auth";
import { audit, recentCount } from "@/lib/audit";

export const runtime = "nodejs";
const DEFAULT_MSG = "EMERGENCY: Reply SAFE, NOT SAFE, or HELP to 6780 immediately.";
const MAX_PER_10_MIN = 5;

type Target = { type: "ALL" | "SITE" | "DEPT" | "INDIVIDUALS"; siteId?: string; deptId?: string; employeeIds?: string[] };

// Operators may only re-send to people still pending; new broadcasts need Super Admin or Site Commander.
// Everyone except Super Admin is limited to their own site.
export const POST = guard(async (req: NextRequest) => {
  const sess = await requireUser();
  const scope = siteScope(sess);
  const ip = clientIp(req);
  const { title, message = DEFAULT_MSG, target, reblastCampaignId } =
    (await req.json()) as { title?: string; message?: string; target?: Target; reblastCampaignId?: string };

  if (!reblastCampaignId && sess.role === "OPERATOR")
    return NextResponse.json({ error: "Operators can re-send to people who haven't replied, but can't start a new broadcast." }, { status: 403 });

  const action = reblastCampaignId ? "broadcast.reblast" : "broadcast.create";
  // A stolen or misused account can't text the whole company repeatedly.
  if ((await recentCount(sess.id, action, 10)) >= MAX_PER_10_MIN)
    return NextResponse.json({ error: "Too many broadcasts in the last 10 minutes. Wait a few minutes or ask a Super Admin." }, { status: 429 });

  let campaign;
  let employees;
  let fingerprint = "";
  if (reblastCampaignId) {
    campaign = await prisma.broadcastCampaign.findUnique({ where: { id: reblastCampaignId } });
    if (!campaign || (scope && campaign.siteId && campaign.siteId !== scope)) return NextResponse.json({ error: "Broadcast not found." }, { status: 404 });
    employees = await prisma.employee.findMany({
      where: { ...(scope ? { siteId: scope } : {}), responses: { some: { campaignId: campaign.id, status: "PENDING" } } },
      include: { site: true },
    });
  } else {
    if (!target || typeof title !== "string" || !title.trim() || title.length > 120)
      return NextResponse.json({ error: "A title (up to 120 characters) and a target are required." }, { status: 400 });
    if (typeof message !== "string" || !message.trim() || message.length > 320)
      return NextResponse.json({ error: "The message must be 1 to 320 characters." }, { status: 400 });
    fingerprint = createFingerprint(title, message, scope ?? target.type + (target.siteId ?? "") + (target.deptId ?? ""));
    // Blocks a double-click or double-submit from sending twice.
    const dup = await prisma.auditLog.count({ where: { userId: sess.id, action, at: { gte: new Date(Date.now() - 60_000) }, detail: { path: ["fingerprint"], equals: fingerprint } } });
    if (dup) return NextResponse.json({ error: "This exact broadcast was just sent. Wait a minute before sending it again." }, { status: 409 });

    const where =
      scope ? { siteId: scope, ...(target.type === "DEPT" ? { deptId: target.deptId } : target.type === "INDIVIDUALS" ? { id: { in: (target.employeeIds ?? []).slice(0, 5000) } } : {}) } :
      target.type === "SITE" ? { siteId: target.siteId } :
      target.type === "DEPT" ? { deptId: target.deptId } :
      target.type === "INDIVIDUALS" ? { id: { in: (target.employeeIds ?? []).slice(0, 5000) } } : {};
    employees = await prisma.employee.findMany({ where, include: { site: true } });
    if (!employees.length) return NextResponse.json({ error: "No recipients match that target." }, { status: 400 });

    campaign = await prisma.broadcastCampaign.create({ data: { title: title.trim(), messageBody: message.trim(), siteId: scope ?? (target.type === "SITE" ? target.siteId : null) } });
    await prisma.sMSResponse.createMany({ data: employees.map((e) => ({ campaignId: campaign!.id, employeeId: e.id })) });
  }

  // Logged before sending, so even a crash mid-send leaves a record of who triggered it.
  await audit(action, { userId: sess.id, email: sess.email, ip, detail: { campaignId: campaign.id, recipients: employees.length, fingerprint } });

  const bySender = new Map<string, string[]>();
  for (const e of employees) bySender.set(e.site.senderId, [...(bySender.get(e.site.senderId) ?? []), e.phoneNumber]);
  const groups = await Promise.all(
    [...bySender].map(async ([senderAddress, recipients]) => ({
      senderAddress,
      results: await sendSms({ senderAddress, shortCode: campaign!.shortCode, message: campaign!.messageBody, recipients }),
    }))
  );
  let failed = 0;
  for (const g of groups) for (const r of g.results) if (!r.ok) { failed += r.count; console.error("globe send failed", g.senderAddress, r.error); }

  publish({ type: "broadcast", campaignId: campaign.id, at: new Date().toISOString() });
  // Upstream error text stays in server logs; the browser only gets counts.
  return NextResponse.json({ campaignId: campaign.id, recipients: employees.length, failed });
});

function createFingerprint(...parts: string[]) {
  let h = 5381; for (const ch of parts.join("|")) h = ((h << 5) + h + ch.charCodeAt(0)) | 0;
  return String(h >>> 0);
}
