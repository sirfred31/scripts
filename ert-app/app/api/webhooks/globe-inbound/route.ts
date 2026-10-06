import { NextRequest, NextResponse } from "next/server";
import { createHash, timingSafeEqual } from "crypto";
import { prisma } from "@/lib/db";
import { publish } from "@/lib/events";
import { normalizePhone, parseInbound, parseKeyword } from "@/lib/globe";

export const runtime = "nodejs";

const sha = (s: string) => createHash("sha256").update(s).digest();
/** Constant-time compare; fails closed when the secret is missing or weak. */
function secretOk(given: string | null, expected: string | undefined) {
  return !!given && !!expected && expected.length >= 16 && timingSafeEqual(sha(given), sha(expected));
}

export async function POST(req: NextRequest) {
  if (!secretOk(req.headers.get("x-webhook-secret"), process.env.GLOBE_WEBHOOK_SECRET)) {
    console.warn("globe webhook rejected: bad or missing secret");
    return NextResponse.json({ error: "unauthorized" }, { status: 401 });
  }
  const allow = process.env.GLOBE_WEBHOOK_IPS?.split(",").map((x) => x.trim()).filter(Boolean);
  const ip = req.headers.get("x-forwarded-for")?.split(",")[0].trim() ?? "";
  if (allow?.length && !allow.includes(ip)) return NextResponse.json({ error: "unauthorized" }, { status: 401 });

  const text = await req.text();
  if (text.length > 10_000) return NextResponse.json({ error: "too large" }, { status: 413 });
  let json: unknown; try { json = JSON.parse(text); } catch { json = null; }
  const payload = parseInbound(json);
  if (!payload) return NextResponse.json({ error: "bad payload" }, { status: 400 });
  if (payload.shortcode && payload.shortcode !== (process.env.GLOBE_SHORTCODE ?? "6780"))
    return NextResponse.json({ ok: true, matched: false }); // not our shortcode: ignore

  const phone = normalizePhone(payload.sender);
  const employee = phone ? await prisma.employee.findUnique({ where: { phoneNumber: phone } }) : null;
  if (!employee) return NextResponse.json({ ok: true, matched: false }); // 200 so Globe doesn't retry

  const response = await prisma.sMSResponse.findFirst({ where: { employeeId: employee.id }, orderBy: { campaign: { createdAt: "desc" } } });
  if (!response) return NextResponse.json({ ok: true, matched: false });

  const status = parseKeyword(payload.message);
  const rawReply = payload.message.slice(0, 500);
  await prisma.sMSResponse.update({ where: { id: response.id }, data: { rawReply, ...(status ? { status } : {}) } });

  publish({
    type: "reply", campaignId: response.campaignId, employeeId: employee.id,
    name: `${employee.firstName} ${employee.lastName}`, siteId: employee.siteId, phone: employee.phoneNumber,
    status: status ?? "UNRECOGNIZED", rawReply, at: new Date().toISOString(), // server time: don't trust sender timestamps
  });
  return NextResponse.json({ ok: true, matched: true, status: status ?? "UNRECOGNIZED" });
}
