import { guard, requireUser, siteScope } from "@/lib/auth";
import { subscribe } from "@/lib/events";
export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export const GET = guard(async () => {
  const scope = siteScope(await requireUser());
  const enc = new TextEncoder();
  let unsub = () => {};
  let ping: ReturnType<typeof setInterval>;
  const stream = new ReadableStream({
    start(c) {
      // Scoped users only receive replies from their own site.
      unsub = subscribe((e) => { if (!scope || e.type === "broadcast" || e.siteId === scope) c.enqueue(enc.encode(`data: ${JSON.stringify(e)}\n\n`)); });
      ping = setInterval(() => c.enqueue(enc.encode(": ping\n\n")), 25000);
    },
    cancel() { unsub(); clearInterval(ping); },
  });
  return new Response(stream, { headers: { "Content-Type": "text/event-stream", "Cache-Control": "no-cache, no-transform", Connection: "keep-alive" } });
});
