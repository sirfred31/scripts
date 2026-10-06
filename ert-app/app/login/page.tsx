"use client";
import { useState } from "react";

export default function Login() {
  const [f, setF] = useState({ email: "", password: "" });
  const [ticket, setTicket] = useState("");
  const [code, setCode] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  async function post(url: string, body: object) {
    setBusy(true); setErr("");
    const r = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    const j = await r.json().catch(() => ({}));
    setBusy(false);
    if (!r.ok) { setErr(j.error ?? "Could not sign in."); if (r.status === 401 && ticket && /expired/i.test(j.error ?? "")) { setTicket(""); setCode(""); } return; }
    if (j.mfaRequired) { setTicket(j.ticket); return; }
    location.href = "/dashboard";
  }
  const submit = () => (ticket ? post("/api/auth/mfa/verify", { ticket, code }) : post("/api/auth/login", f));

  const field = "w-full rounded-md px-3 py-2 bg-transparent ring-1 ring-slate-700 focus-visible:outline focus-visible:outline-2 focus-visible:outline-slate-300";
  return (
    <main className="min-h-screen grid place-items-center bg-slate-950 text-slate-100 p-4">
      <div className="w-full max-w-sm space-y-4" onKeyDown={(e) => e.key === "Enter" && !busy && submit()}>
        <h1 className="text-xl font-semibold">{ticket ? "Enter your code" : "Emergency response sign in"}</h1>
        {err && <p role="alert" className="rounded-md bg-red-950 text-red-200 px-3 py-2 text-sm">{err}</p>}
        {ticket ? (
          <>
            <label className="block text-sm">6-digit code from your authenticator app, or a recovery code
              <input autoFocus inputMode="text" autoComplete="one-time-code" className={`${field} mt-1 tracking-widest`} value={code} onChange={(e) => setCode(e.target.value)} />
            </label>
            <button className="text-sm underline" onClick={() => { setTicket(""); setCode(""); setErr(""); }}>Start over</button>
          </>
        ) : (
          <>
            <label className="block text-sm">Email<input type="email" autoComplete="username" className={`${field} mt-1`} value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} /></label>
            <label className="block text-sm">Password<input type="password" autoComplete="current-password" className={`${field} mt-1`} value={f.password} onChange={(e) => setF({ ...f, password: e.target.value })} /></label>
          </>
        )}
        <button disabled={busy || (ticket ? !code : !f.email || !f.password)} onClick={submit} className="w-full rounded-md bg-slate-100 text-slate-900 font-semibold py-2.5 disabled:opacity-40">
          {busy ? "Checking…" : ticket ? "Verify" : "Sign in"}
        </button>
      </div>
    </main>
  );
}
