"use client";
import { useCallback, useEffect, useState } from "react";
import { Phone, RefreshCw, Siren, Moon, Sun, X } from "lucide-react";

type Row = {
  id: string; status: "PENDING" | "SAFE" | "NOT_SAFE" | "HELP"; rawReply: string | null; updatedAt: string;
  employee: { firstName: string; lastName: string; phoneNumber: string; site: { name: string }; department: { name: string } };
};
type Data = {
  campaign: { id: string; title: string; messageBody: string } | null;
  sites: { id: string; name: string; senderId: string }[];
  counts: { total: number; safe: number; notSafe: number; help: number; pending: number } | null;
  feed: Row[]; unresponsive: Row[];
};

const BADGE: Record<string, string> = {
  SAFE: "bg-emerald-500/15 text-emerald-600 dark:text-emerald-400 ring-emerald-500/40",
  NOT_SAFE: "bg-amber-500/15 text-amber-600 dark:text-amber-400 ring-amber-500/40",
  HELP: "bg-red-600 text-white ring-red-400 animate-pulse",
};
const LABEL: Record<string, string> = { SAFE: "Safe", NOT_SAFE: "Not safe", HELP: "Help" };

export default function Dashboard() {
  const [dark, setDark] = useState(true);
  const [siteId, setSiteId] = useState("");
  const [data, setData] = useState<Data | null>(null);
  const [modal, setModal] = useState<0 | 1 | 2>(0); // 0 closed, 1 compose, 2 final confirm
  const [form, setForm] = useState({ title: "", message: "EMERGENCY: Reply SAFE, NOT SAFE, or HELP to 6780 immediately.", scope: "ALL", siteId: "" });
  const [typed, setTyped] = useState("");
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const [me, setMe] = useState<{ email: string; role: string; siteId: string | null } | null>(null);

  const load = useCallback(async () => {
    const r = await fetch(`/api/dashboard${siteId ? `?siteId=${siteId}` : ""}`);
    if (r.status === 401) { location.href = "/login"; return; }
    if (r.status === 403) { location.href = "/security"; return; }
    setData(await r.json());
  }, [siteId]);

  useEffect(() => { load(); }, [load]);
  useEffect(() => { fetch("/api/auth/me").then((r) => r.json()).then((m) => { if (m.role === "SUPER_ADMIN" && !m.mfaEnabled) location.href = "/security"; setMe(m); }); }, []);
  useEffect(() => { // live updates: refetch whenever the server pushes a reply or broadcast
    const es = new EventSource("/api/events");
    es.onmessage = () => load();
    return () => es.close();
  }, [load]);

  async function send(body: object) {
    setBusy(true);
    const r = await fetch("/api/broadcast", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    const j = await r.json();
    setBusy(false); setModal(0); setTyped("");
    setNotice(r.ok ? (j.failed ? `Sent to ${j.recipients - j.failed} of ${j.recipients} people. ${j.failed} could not be sent. Check the Globe connection and send again.` : `Sent to ${j.recipients} people.`) : `Broadcast not sent: ${j.error ?? r.status}`);
    load();
  }
  const launch = () => send({
    title: form.title, message: form.message,
    target: me?.siteId ? { type: "SITE", siteId: me.siteId } : form.scope === "SITE" ? { type: "SITE", siteId: form.siteId } : { type: "ALL" },
  });

  const c = data?.counts;
  const metrics = [
    ["Sent", c?.total, ""], ["Safe", c?.safe, "text-emerald-500"], ["Not safe", c?.notSafe, "text-amber-500"],
    ["Help requested", c?.help, "text-red-500"], ["No reply", c?.pending, "text-slate-400"],
  ] as const;

  return (
    <div className={dark ? "dark" : ""}>
      <main className="min-h-screen bg-slate-50 text-slate-900 dark:bg-slate-950 dark:text-slate-100 p-4 md:p-6">
        <header className="flex flex-wrap items-center gap-3 mb-5">
          <div className="mr-auto">
            <h1 className="text-xl font-semibold">Emergency response</h1>
            <p className="text-sm text-slate-500">{data?.campaign ? data.campaign.title : "No broadcast sent yet"}</p>
          </div>
          <a href="/directory" className="text-sm underline">Directory</a>
          {me?.role === "SUPER_ADMIN" && <a href="/users" className="text-sm underline">Users</a>}
          <a href="/security" className="text-sm underline">Security</a>
          <button onClick={async () => { await fetch("/api/auth/logout", { method: "POST" }); location.href = "/login"; }} className="text-sm underline">Sign out</button>
          <button onClick={() => setDark(!dark)} aria-label="Toggle dark mode" className="p-2 rounded-md ring-1 ring-slate-300 dark:ring-slate-700">
            {dark ? <Sun size={18} /> : <Moon size={18} />}
          </button>
          <button onClick={() => setModal(1)} className={`${me?.role === "OPERATOR" ? "hidden" : "flex"} items-center gap-2 rounded-md bg-red-600 hover:bg-red-700 text-white font-semibold px-4 py-2.5 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-red-400`}>
            <Siren size={18} /> INITIATE EMERGENCY BROADCAST
          </button>
        </header>

        {notice && <p role="status" className="mb-4 text-sm rounded-md bg-slate-200 dark:bg-slate-800 px-3 py-2">{notice}</p>}

        <nav className="flex gap-2 mb-5" aria-label="Site filter">
          {[...(me?.siteId ? [] : [{ id: "", name: "All sites" }]), ...(data?.sites ?? [])].map((s) => (
            <button key={s.id} onClick={() => setSiteId(s.id)}
              className={`px-3 py-1.5 rounded-md text-sm ring-1 ${siteId === s.id ? "bg-slate-900 text-white dark:bg-slate-100 dark:text-slate-900 ring-transparent" : "ring-slate-300 dark:ring-slate-700"}`}>
              {s.name}
            </button>
          ))}
        </nav>

        <section className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-6">
          {metrics.map(([label, v, color]) => (
            <div key={label} className="rounded-lg p-4 bg-white dark:bg-slate-900 ring-1 ring-slate-200 dark:ring-slate-800">
              <div className={`text-3xl font-bold tabular-nums ${color}`}>{v ?? "–"}</div>
              <div className="text-sm text-slate-500">{label}</div>
            </div>
          ))}
        </section>

        <div className="grid lg:grid-cols-3 gap-6">
          <section className="lg:col-span-2">
            <h2 className="font-semibold mb-2">Replies</h2>
            <ul className="space-y-2">
              {data?.feed.length === 0 && <li className="text-sm text-slate-500">No replies yet. They appear here as employees text 6780.</li>}
              {data?.feed.map((r) => (
                <li key={r.id} className={`flex items-center gap-3 rounded-lg p-3 bg-white dark:bg-slate-900 ring-1 ${r.status === "HELP" ? "ring-red-500" : "ring-slate-200 dark:ring-slate-800"}`}>
                  <span className={`px-2.5 py-1 rounded-full text-xs font-semibold ring-1 ${BADGE[r.status]}`}>{LABEL[r.status]}</span>
                  <div className="mr-auto min-w-0">
                    <div className="font-medium truncate">{r.employee.firstName} {r.employee.lastName}</div>
                    <div className="text-xs text-slate-500">{r.employee.site.name}, {r.employee.department.name}</div>
                  </div>
                  <time className="text-xs text-slate-500">{new Date(r.updatedAt).toLocaleTimeString()}</time>
                  {r.status !== "SAFE" && <a href={`tel:+${r.employee.phoneNumber}`} aria-label={`Call ${r.employee.firstName}`} className="p-2 rounded-md bg-slate-100 dark:bg-slate-800"><Phone size={16} /></a>}
                </li>
              ))}
            </ul>
          </section>

          <aside>
            <div className="flex items-center mb-2">
              <h2 className="font-semibold mr-auto">No reply ({data?.unresponsive.length ?? 0})</h2>
              <button disabled={!data?.unresponsive.length || busy}
                onClick={() => data?.campaign && send({ reblastCampaignId: data.campaign.id })}
                className="flex items-center gap-1.5 text-sm rounded-md px-3 py-1.5 ring-1 ring-slate-300 dark:ring-slate-700 disabled:opacity-40">
                <RefreshCw size={14} /> Send again
              </button>
            </div>
            <ul className="space-y-2 max-h-[60vh] overflow-y-auto">
              {data?.unresponsive.map((r) => (
                <li key={r.id} className="flex items-center gap-3 rounded-lg p-3 bg-white dark:bg-slate-900 ring-1 ring-slate-200 dark:ring-slate-800">
                  <div className="mr-auto min-w-0">
                    <div className="truncate">{r.employee.firstName} {r.employee.lastName}</div>
                    <div className="text-xs text-slate-500">{r.employee.site.name}</div>
                  </div>
                  <a href={`tel:+${r.employee.phoneNumber}`} aria-label={`Call ${r.employee.firstName}`} className="p-2 rounded-md bg-slate-100 dark:bg-slate-800"><Phone size={16} /></a>
                </li>
              ))}
            </ul>
          </aside>
        </div>

        {modal > 0 && (
          <div role="dialog" aria-modal="true" className="fixed inset-0 z-50 grid place-items-center bg-black/70 p-4">
            <div className="w-full max-w-lg rounded-xl bg-white dark:bg-slate-900 p-6 ring-1 ring-slate-700">
              <div className="flex items-center mb-4">
                <h2 className="text-lg font-semibold mr-auto">{modal === 1 ? "Compose broadcast" : "Confirm broadcast"}</h2>
                <button onClick={() => { setModal(0); setTyped(""); }} aria-label="Close"><X size={18} /></button>
              </div>
              {modal === 1 ? (
                <div className="space-y-3">
                  <input value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} placeholder="Title, e.g. Earthquake alert Oct 2026" className="w-full rounded-md px-3 py-2 bg-transparent ring-1 ring-slate-300 dark:ring-slate-700" />
                  <textarea value={form.message} onChange={(e) => setForm({ ...form, message: e.target.value })} rows={3} className="w-full rounded-md px-3 py-2 bg-transparent ring-1 ring-slate-300 dark:ring-slate-700" />
                  <div className={me?.siteId ? "hidden" : "flex gap-2"}>
                    <select value={form.scope} onChange={(e) => setForm({ ...form, scope: e.target.value })} className="rounded-md px-3 py-2 bg-transparent ring-1 ring-slate-300 dark:ring-slate-700">
                      <option value="ALL">All sites</option><option value="SITE">One site</option>
                    </select>
                    {form.scope === "SITE" && (
                      <select value={form.siteId} onChange={(e) => setForm({ ...form, siteId: e.target.value })} className="rounded-md px-3 py-2 bg-transparent ring-1 ring-slate-300 dark:ring-slate-700">
                        <option value="">Choose site</option>
                        {data?.sites.map((s) => <option key={s.id} value={s.id}>{s.name} ({s.senderId})</option>)}
                      </select>
                    )}
                  </div>
                  <button disabled={!form.title || (form.scope === "SITE" && !form.siteId)} onClick={() => setModal(2)} className="w-full rounded-md bg-red-600 text-white font-semibold py-2.5 disabled:opacity-40">Review</button>
                </div>
              ) : (
                <div className="space-y-3">
                  <p className="text-sm">This texts {form.scope === "ALL" && !me?.siteId ? "every employee at all sites" : `every employee at ${data?.sites.find((s) => s.id === (me?.siteId ?? form.siteId))?.name}`} right away:</p>
                  <blockquote className="rounded-md bg-slate-100 dark:bg-slate-800 p-3 text-sm">{form.message}</blockquote>
                  <label className="block text-sm">Type SEND to confirm
                    <input value={typed} onChange={(e) => setTyped(e.target.value)} className="mt-1 w-full rounded-md px-3 py-2 bg-transparent ring-1 ring-slate-300 dark:ring-slate-700" />
                  </label>
                  <div className="flex gap-2">
                    <button onClick={() => setModal(1)} className="flex-1 rounded-md py-2.5 ring-1 ring-slate-300 dark:ring-slate-700">Back</button>
                    <button disabled={typed.trim().toUpperCase() !== "SEND" || busy} onClick={launch} className="flex-1 rounded-md bg-red-600 text-white font-semibold py-2.5 disabled:opacity-40">{busy ? "Sending…" : "Send broadcast"}</button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
