"use client";
import { useEffect, useState } from "react";

type Me = { email: string; role: string; mfaEnabled: boolean };
const field = "rounded-md px-3 py-2 bg-transparent ring-1 ring-slate-700 focus-visible:outline focus-visible:outline-2 focus-visible:outline-slate-300";

export default function Security() {
  const [me, setMe] = useState<Me | null>(null);
  const [setup, setSetup] = useState<{ secret: string; qr: string } | null>(null);
  const [code, setCode] = useState("");
  const [codes, setCodes] = useState<string[] | null>(null);
  const [err, setErr] = useState("");

  const load = () => fetch("/api/auth/me").then((r) => (r.status === 401 ? (location.href = "/login", null) : r.json())).then((m) => m && setMe(m));
  useEffect(() => { load(); }, []);

  async function post(url: string, body?: object) {
    const r = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body ?? {}) });
    const j = await r.json(); setErr(r.ok ? "" : j.error ?? "Something went wrong.");
    return r.ok ? j : null;
  }

  const required = me?.role === "SUPER_ADMIN";
  return (
    <main className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-6">
      <div className="max-w-lg mx-auto space-y-5">
        <header className="flex items-center gap-3">
          <h1 className="text-xl font-semibold mr-auto">Two-step sign-in</h1>
          {me?.mfaEnabled && <a href="/dashboard" className="text-sm underline">Back to dashboard</a>}
          <button className="text-sm underline" onClick={async () => { await fetch("/api/auth/logout", { method: "POST" }); location.href = "/login"; }}>Sign out</button>
        </header>
        {err && <p role="alert" className="rounded-md bg-red-950 text-red-200 px-3 py-2 text-sm">{err}</p>}

        {codes ? (
          <section className="space-y-3">
            <p className="rounded-md bg-emerald-950 text-emerald-200 px-3 py-2 text-sm">Two-step sign-in is on.</p>
            <h2 className="font-semibold">Save your recovery codes</h2>
            <p className="text-sm text-slate-400">Each code works once if you lose your phone. They won't be shown again. Store them somewhere safe, apart from your password.</p>
            <ul className="grid grid-cols-2 gap-2 font-mono text-lg">{codes.map((c) => <li key={c} className="rounded-md bg-slate-900 ring-1 ring-slate-800 px-3 py-2">{c}</li>)}</ul>
            <button className="rounded-md bg-slate-100 text-slate-900 font-medium px-4 py-2" onClick={() => { setCodes(null); load(); }}>I saved them</button>
          </section>
        ) : me?.mfaEnabled ? (
          <section className="space-y-3">
            <p className="rounded-md bg-emerald-950 text-emerald-200 px-3 py-2 text-sm">Two-step sign-in is on for {me.email}.</p>
            {required ? <p className="text-sm text-slate-400">Super Admins must keep it on. If you lose your phone and recovery codes, another Super Admin can reset it for you on the Users page.</p> : (
              <div className="space-y-2">
                <p className="text-sm text-slate-400">To turn it off, enter a current code from your app.</p>
                <div className="flex gap-2">
                  <input className={`${field} w-40`} inputMode="numeric" placeholder="6-digit code" value={code} onChange={(e) => setCode(e.target.value)} />
                  <button className={field} onClick={async () => { if (await post("/api/auth/mfa/disable", { code })) { setCode(""); load(); } }}>Turn off</button>
                </div>
              </div>
            )}
          </section>
        ) : setup ? (
          <section className="space-y-3">
            <p className="text-sm text-slate-400">1. Open an authenticator app (Google Authenticator, Microsoft Authenticator, Authy, 1Password) and scan this code.</p>
            {/* eslint-disable-next-line @next/next/no-img-element */}
            <img src={setup.qr} alt="QR code for your authenticator app" width={220} height={220} className="rounded-md bg-white p-2" />
            <p className="text-sm text-slate-400">Can't scan? Enter this key by hand: <code className="select-all text-slate-100">{setup.secret}</code></p>
            <p className="text-sm text-slate-400">2. Type the 6-digit code the app shows.</p>
            <div className="flex gap-2">
              <input autoFocus className={`${field} w-40`} inputMode="numeric" autoComplete="one-time-code" placeholder="6-digit code" value={code} onChange={(e) => setCode(e.target.value)} />
              <button className="rounded-md bg-slate-100 text-slate-900 font-medium px-4" onClick={async () => { const j = await post("/api/auth/mfa/enable", { code }); if (j) { setCodes(j.recoveryCodes); setSetup(null); setCode(""); } }}>Turn on</button>
            </div>
          </section>
        ) : (
          <section className="space-y-3">
            {required && <p className="rounded-md bg-amber-950 text-amber-200 px-3 py-2 text-sm">Super Admin accounts need two-step sign-in before you can continue.</p>}
            <p className="text-sm text-slate-400">After your password, you'll enter a 6-digit code from an app on your phone. A stolen password alone then can't send an emergency broadcast.</p>
            <button className="rounded-md bg-slate-100 text-slate-900 font-medium px-4 py-2" onClick={async () => { const j = await post("/api/auth/mfa/setup"); if (j) setSetup(j); }}>Set up two-step sign-in</button>
          </section>
        )}
      </div>
    </main>
  );
}
