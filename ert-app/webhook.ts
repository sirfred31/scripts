import { createHash, timingSafeEqual } from "crypto";
import type { NextRequest } from "next/server";

const sha = (s: string) => createHash("sha256").update(s).digest();

/** Shared by both Globe webhooks. Constant-time secret check that fails closed, plus an optional IP allow-list. */
export function webhookAuthorized(req: NextRequest) {
  const given = req.headers.get("x-webhook-secret"), expected = process.env.GLOBE_WEBHOOK_SECRET;
  if (!given || !expected || expected.length < 16 || !timingSafeEqual(sha(given), sha(expected))) return false;
  const allow = process.env.GLOBE_WEBHOOK_IPS?.split(",").map((x) => x.trim()).filter(Boolean);
  const ip = req.headers.get("x-forwarded-for")?.split(",")[0].trim() ?? "";
  return !allow?.length || allow.includes(ip);
}
