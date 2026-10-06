// In-process pub/sub feeding the SSE endpoint.
// Single-instance only: on multiple instances/serverless, swap for Redis pub/sub or Supabase Realtime.
export type ErtEvent = {
  type: "reply" | "broadcast";
  campaignId: string;
  employeeId?: string;
  name?: string;
  siteId?: string;
  phone?: string;
  status?: "SAFE" | "NOT_SAFE" | "HELP" | "UNRECOGNIZED";
  rawReply?: string;
  at: string;
};
type Listener = (e: ErtEvent) => void;
const g = globalThis as unknown as { ertListeners?: Set<Listener> };
const listeners = (g.ertListeners ??= new Set<Listener>());
export const subscribe = (l: Listener) => (listeners.add(l), () => listeners.delete(l));
export const publish = (e: ErtEvent) => listeners.forEach((l) => l(e));
