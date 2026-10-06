"use client";
import { useCallback, useEffect, useState } from "react";

type User = { id: string; email: string; role: string; active: boolean; mfaEnabled: boolean; siteId: string | null; site: { name: string } | null };
type Site = { id: string; name: string };
type Log = { id: string; at: string; email: string | null; action: string; ip: string | null; detail: Record<string, unknown> };

const input = "rounded-md px-3 py-2 bg-transparent ring-1 ring-slate-700 focus-visible:outline focus-visible:outline-2 focus-visible:outline-slate-300";
const ROLE: Record<string, string> = { SUPER_ADMIN: "Super Admin", SITE_COMMANDER: "Site Commander", OPERATOR: "Operator" };

export default function Users() {
  const [users, setUsers] = useState<User[]>([]);
  const [sites, setSites] = useState<Site[]>([]);
  const [log, setLog] = useState<Log[]>([]);
  const [err, setErr] = useState("");
  const [f, setF] = useState({ email: "", password: "", role: "OPERATOR", siteId: "" });
  const [reset, setReset] = useState<{ id: string; pw: string } | null>(null);

  const load = useCallback(async () => {
    const [u, s, a] = await Promise.all([fetch("/api/users"), fetch("/api/sites"), fetch("/api/audit")]);
    if (u.status === 401) { location.href = "/login"; return; }
    if (u.status === 403) { const j = await u.json(); if (j.code === "MFA_SETUP_REQUIRED") { location.href = "/security"; return; } setErr("Only Super Admins can manage users."); return; }
    setUsers(await u.json()); setSites(await s.json()); setLog(await a.json());
  }, []);
  useEffect(() => { load(); }, [load]);

  async function call(url: string, method: string, body: unknown) {
    const r = await fetch(url, { method, headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    const j = await r.json();
    setErr(r.ok ? "" : j.error ?? "Something went wrong.");
    if (r.ok) load();
    return r.ok;
  }

  return (
    <div className="dark">
      <main className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-6 space-y-8">
        <header className="flex items-center gap-3">
          <h1 className="text-xl font-semibold mr-auto">Users</h1>
          <a href="/dashboard" className="text-sm underline">Back to dashboard</a>
        </header>
        {err && <p role="alert" className="rounded-md bg-red-950 text-red-200 px-3 py-2 text-sm">{err}</p>}

        <section className="space-y-3">
          <h2 className="font-semibold">Add a user</h2>
          <div className="flex flex-wrap gap-2">
            <input className={`${input} w-64`} type="email" placeholder="Email" autoComplete="off" value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} />
            <input className={`${input} w-56`} type="password" placeholder="Temporary password (12+ characters)" autoComplete="new-password" value={f.password} onChange={(e) => setF({ ...f, password: e.target.value })} />
            <select className={input} value={f.role} onChange={(e) => setF({ ...f, role: e.target.value })}>
              {Object.entries(ROLE).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
            </select>
            {f.role !== "SUPER_ADMIN" && (
              <select className={input} value={f.siteId} onChange={(e) => setF({ ...f, siteId: e.target.value })}>
                <option value="">Site</option>{sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            )}
            <button className="rounded-md bg-slate-100 text-slate-900 font-medium px-4" onClick={async () => (await call("/api/users", "POST", f)) && setF({ email: "", password: "", role: "OPERATOR", siteId: "" })}>Add user</button>
          </div>
        </section>

        <section className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="text-left text-slate-500"><tr><th className="py-2">Email</th><th>Role</th><th>Site</th><th>Two-step</th><th>Status</th><th /></tr></thead>
            <tbody>
              {users.map((u) => (
                <tr key={u.id} className="border-t border-slate-800">
                  <td className="py-2">{u.email}</td><td>{ROLE[u.role]}</td><td>{u.site?.name ?? "All sites"}</td>
                  <td className={u.mfaEnabled ? "text-emerald-400" : "text-amber-400"}>{u.mfaEnabled ? "On" : "Off"}</td>
                  <td className={u.active ? "text-emerald-400" : "text-slate-500"}>{u.active ? "Active" : "Deactivated"}</td>
                  <td className="text-right space-x-3">
                    {reset?.id === u.id ? (
                      <span className="inline-flex gap-2">
                        <input className={`${input} w-52 py-1`} type="password" placeholder="New password" autoComplete="new-password" value={reset.pw} onChange={(e) => setReset({ id: u.id, pw: e.target.value })} />
                        <button className="underline" onClick={async () => (await call(`/api/users/${u.id}`, "PATCH", { password: reset.pw })) && setReset(null)}>Save</button>
                        <button className="underline" onClick={() => setReset(null)}>Cancel</button>
                      </span>
                    ) : (
                      <>
                        <button className="underline" onClick={() => setReset({ id: u.id, pw: "" })}>Reset password</button>
                        {u.mfaEnabled && <button className="underline" onClick={() => confirm(`Reset two-step sign-in for ${u.email}? They will set it up again at next sign-in.`) && call(`/api/users/${u.id}`, "PATCH", { resetMfa: true })}>Reset two-step</button>}
                        <button className="underline" onClick={() => call(`/api/users/${u.id}`, "PATCH", { active: !u.active })}>{u.active ? "Deactivate" : "Reactivate"}</button>
                      </>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="mt-2 text-xs text-slate-500">Resetting a password or deactivating a user signs them out everywhere right away.</p>
        </section>

        <section className="space-y-2">
          <h2 className="font-semibold">Recent activity</h2>
          <ul className="text-sm space-y-1 max-h-96 overflow-y-auto">
            {log.map((l) => (
              <li key={l.id} className={`flex gap-3 ${l.action === "login.fail" ? "text-amber-400" : ""}`}>
                <time className="text-slate-500 tabular-nums shrink-0">{new Date(l.at).toLocaleString()}</time>
                <span className="shrink-0">{l.action}</span>
                <span className="text-slate-400 truncate">{l.email ?? "unknown"}{l.ip ? `, ${l.ip}` : ""}{Object.keys(l.detail).length ? `, ${JSON.stringify(l.detail)}` : ""}</span>
              </li>
            ))}
          </ul>
        </section>
      </main>
    </div>
  );
}
