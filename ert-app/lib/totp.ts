import { createCipheriv, createDecipheriv, createHash, createHmac, randomBytes, timingSafeEqual } from "crypto";
import bcrypt from "bcryptjs";

const B32 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
export function toBase32(buf: Buffer) {
  let bits = 0, val = 0, out = "";
  for (const b of buf) {
    val = (val << 8) | b; bits += 8;
    while (bits >= 5) { out += B32[(val >>> (bits - 5)) & 31]; bits -= 5; }
    val &= (1 << bits) - 1;
  }
  return bits ? out + B32[(val << (5 - bits)) & 31] : out;
}
function fromBase32(s: string) {
  let bits = 0, val = 0; const out: number[] = [];
  for (const c of s.toUpperCase()) {
    const i = B32.indexOf(c); if (i < 0) throw new Error("bad base32");
    val = (val << 5) | i; bits += 5;
    if (bits >= 8) { out.push((val >>> (bits - 8)) & 255); bits -= 8; }
    val &= (1 << bits) - 1;
  }
  return Buffer.from(out);
}

export const newSecret = () => toBase32(randomBytes(20));

function hotp(secret: Buffer, counter: number) {
  const msg = Buffer.alloc(8); msg.writeBigUInt64BE(BigInt(counter));
  const h = createHmac("sha1", secret).update(msg).digest();
  const o = h[19] & 15;
  const n = (((h[o] & 0x7f) << 24) | (h[o + 1] << 16) | (h[o + 2] << 8) | h[o + 3]) % 1_000_000;
  return String(n).padStart(6, "0");
}

/** Accepts the current step and one either side (clock drift). Returns the matched step, or null.
 *  Steps at or below lastStep are refused, so a captured code can't be used twice. */
export function verifyTotp(secretB32: string, code: string, lastStep: number) {
  if (!/^\d{6}$/.test(code)) return null;
  const secret = fromBase32(secretB32), now = Math.floor(Date.now() / 30_000);
  let hit: number | null = null;
  for (const step of [now - 1, now, now + 1]) {
    const ok = timingSafeEqual(Buffer.from(hotp(secret, step)), Buffer.from(code));
    if (ok && step > lastStep && hit === null) hit = step; // no early exit: constant work
  }
  return hit;
}

export const otpauthUri = (secret: string, email: string) =>
  `otpauth://totp/${encodeURIComponent("Emergency Response")}:${encodeURIComponent(email)}?secret=${secret}&issuer=${encodeURIComponent("Emergency Response")}&algorithm=SHA1&digits=6&period=30`;

// The TOTP secret is encrypted at rest with its own key, so a database leak alone doesn't expose it.
const encKey = () => {
  const s = process.env.MFA_ENCRYPTION_KEY;
  if (!s || s.length < 32) throw new Error("MFA_ENCRYPTION_KEY must be set to at least 32 characters.");
  return createHash("sha256").update(s).digest();
};
export function encrypt(plain: string) {
  const iv = randomBytes(12), c = createCipheriv("aes-256-gcm", encKey(), iv);
  const data = Buffer.concat([c.update(plain, "utf8"), c.final()]);
  return [iv, c.getAuthTag(), data].map((b) => b.toString("base64")).join(".");
}
export function decrypt(blob: string) {
  const [iv, tag, data] = blob.split(".").map((x) => Buffer.from(x, "base64"));
  const d = createDecipheriv("aes-256-gcm", encKey(), iv); d.setAuthTag(tag);
  return Buffer.concat([d.update(data), d.final()]).toString("utf8");
}

// Recovery codes: 10 random base32 characters, shown once, stored only as bcrypt hashes, single use.
export const newRecoveryCode = () => { const c = toBase32(randomBytes(7)).slice(0, 10); return `${c.slice(0, 5)}-${c.slice(5)}`; };
const norm = (c: string) => c.replace(/[\s-]/g, "").toUpperCase();
export const hashRecovery = (c: string) => bcrypt.hash(norm(c), 10);
export async function matchRecovery(code: string, hashes: string[]) {
  const n = norm(code);
  if (!/^[A-Z2-7]{10}$/.test(n)) return -1;
  let found = -1;
  for (let i = 0; i < hashes.length; i++) if (await bcrypt.compare(n, hashes[i]) && found < 0) found = i;
  return found;
}
