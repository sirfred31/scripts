// In-memory fixed-window limiter. Fine for one instance; use Redis if you run several.
type E = { n: number; reset: number };
const m = new Map<string, E>();
const get = (k: string) => { const e = m.get(k); if (e && e.reset <= Date.now()) { m.delete(k); return undefined; } return e; };
export const isLimited = (key: string, max: number) => (get(key)?.n ?? 0) >= max;
export function hit(key: string, windowMs: number) {
  const e = get(key);
  if (e) { e.n++; return; }
  if (m.size > 10_000) { const now = Date.now(); for (const [k, v] of m) if (v.reset <= now) m.delete(k); if (m.size > 10_000) m.clear(); }
  m.set(key, { n: 1, reset: Date.now() + windowMs });
}
export const clear = (key: string) => void m.delete(key);
