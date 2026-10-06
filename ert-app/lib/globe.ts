import type { Status } from "@prisma/client";

const URL_ = process.env.GLOBE_API_URL!;     // m360 outbound endpoint from your account manager
const TOKEN = process.env.GLOBE_API_TOKEN!;  // adjust auth header to match your m360 credentials
const CHUNK = 100;                           // recipients per request; confirm your account limit

/** Normalize 09XXXXXXXXX, +639XXXXXXXXX, 9XXXXXXXXX -> 639XXXXXXXXX. Returns null if invalid. */
export function normalizePhone(raw: string): string | null {
  const d = raw.replace(/\D/g, "");
  const n = d.startsWith("63") ? d : d.startsWith("0") ? "63" + d.slice(1) : d.length === 10 ? "63" + d : d;
  return /^639\d{9}$/.test(n) ? n : null;
}

export async function sendSms(p: { senderAddress: string; shortCode: string; message: string; recipients: string[] }) {
  const results: { ok: boolean; count: number; error?: string }[] = [];
  for (let i = 0; i < p.recipients.length; i += CHUNK) {
    const recipients = p.recipients.slice(i, i + CHUNK);
    try {
      const res = await fetch(URL_, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${TOKEN}` },
        body: JSON.stringify({ senderAddress: p.senderAddress, shortCode: p.shortCode, message: p.message, recipients }),
      });
      results.push(res.ok ? { ok: true, count: recipients.length } : { ok: false, count: recipients.length, error: `${res.status} ${await res.text()}` });
    } catch (e) {
      results.push({ ok: false, count: recipients.length, error: String(e) });
    }
  }
  return results;
}

export type InboundPayload = { sender: string; shortcode: string; message: string; timestamp?: string };

export function parseInbound(body: unknown): InboundPayload | null {
  const b = body as Partial<InboundPayload>;
  if (!b || typeof b.sender !== "string" || typeof b.message !== "string") return null;
  return { sender: b.sender, shortcode: String(b.shortcode ?? ""), message: b.message, timestamp: b.timestamp };
}

/** Uppercase, trim, collapse spaces. Accepts "NOTSAFE", "not safe", "SAFE." etc. */
export function parseKeyword(text: string): Exclude<Status, "PENDING"> | null {
  const t = text.toUpperCase().replace(/[^A-Z ]/g, " ").replace(/\s+/g, " ").trim();
  if (t === "SAFE") return "SAFE";
  if (t === "NOT SAFE" || t === "NOTSAFE") return "NOT_SAFE";
  if (t === "HELP") return "HELP";
  return null;
}
