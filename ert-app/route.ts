import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { publish } from "@/lib/events";
import { mapDelivery, normalizePhone } from "@/lib/globe";
import { webhookAuthorized } from "@/lib/webhook";

export const runtime = "nodejs";
const RANK = { QUEUED: 0, SENT: 1, FAILED: 2, DELIVERED: 3 } as const; // receipts only ever move forward

type Raw = Record<string, unknown>;
const pick = (o: Raw, ...keys: string[]) => { for (const k of keys) if (o[k] != null && o[k] !== "") return String(o[k]); };

// Delivery receipts from Globe/m360. The parser accepts one report or a list, and common field names.
// Adjust the field names here once you've seen a real receipt from your account.
export async function POST(req: NextRequest) {
  if (!webhookAuthorized(req)) {
    console.warn("globe dlr webhook rejected");
    return NextResponse.json({ error: "unauthorized" }, { status: 401 });
  }
  const text = await req.text();
  if (text.length > 100_000) return NextResponse.json({ error: "too large" }, { status: 413 });
  let body: unknown; try { body = JSON.parse(text); } catch { return NextResponse.json({ error: "bad payload" }, { status: 400 }); }
  const b = body as Raw;
  const items = (Array.isArray(body) ? body : (b.reports ?? b.deliveryReports ?? b.messages ?? [body])) as Raw[];

  let applied = 0;
  for (const it of items.slice(0, 500)) {
    const id = pick(it, "messageId", "message_id", "id", "requestId");
    const phone = normalizePhone((pick(it, "recipient", "to", "msisdn", "destination", "address") ?? "").replace(/^tel:/, ""));
    const next = mapDelivery(pick(it, "status", "deliveryStatus", "state") ?? "");
    if (!next) continue;

    let r = id ? await prisma.sMSResponse.findFirst({ where: { providerMessageId: id }, include: { employee: true } }) : null;
    if (!r && phone) {
      const e = await prisma.employee.findUnique({ where: { phoneNumber: phone } });
      if (e) r = await prisma.sMSResponse.findFirst({ where: { employeeId: e.id }, orderBy: { campaign: { createdAt: "desc" } }, include: { employee: true } });
    }
    if (!r || RANK[next] <= RANK[r.delivery]) continue;

    await prisma.sMSResponse.update({
      where: { id: r.id },
      data: {
        delivery: next,
        deliveredAt: next === "DELIVERED" ? new Date() : undefined,
        deliveryNote: next === "FAILED" ? (pick(it, "reason", "error", "description", "errorCode") ?? "Carrier reported the text as undeliverable").slice(0, 200) : null,
      },
    });
    publish({ type: "delivery", campaignId: r.campaignId, employeeId: r.employeeId, siteId: r.employee.siteId, at: new Date().toISOString() });
    applied++;
  }
  return NextResponse.json({ ok: true, applied });
}
