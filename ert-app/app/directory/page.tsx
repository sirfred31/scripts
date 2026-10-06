"use client";
import { useCallback, useEffect, useRef, useState } from "react";
import { Pencil, Trash2, Upload, Plus } from "lucide-react";

type Site = { id: string; name: string; senderId: string; _count: { employees: number } };
type Dept = { id: string; name: string; _count: { employees: number } };
type Emp = { id: string; firstName: string; lastName: string; phoneNumber: string; siteId: string; deptId: string; site: Site; department: Dept };
type ImportResult = { created: number; updated: number; errors: { row: number; reason: string }[] };

const input = "rounded-md px-3 py-2 bg-transparent ring-1 ring-slate-300 dark:ring-slate-700";
const blank = { id: "", firstName: "", lastName: "", phoneNumber: "", siteId: "", deptId: "" };

export default function Directory() {
  const [emps, setEmps] = useState<Emp[]>([]);
  const [sites, setSites] = useState<Site[]>([]);
  const [depts, setDepts] = useState<Dept[]>([]);
  const [f, setF] = useState({ siteId: "", deptId: "", q: "" });
  const [form, setForm] = useState(blank);
  const [err, setErr] = useState("");
  const [imp, setImp] = useState<ImportResult | null>(null);
  const [newSite, setNewSite] = useState({ name: "", senderId: "" });
  const [newDept, setNewDept] = useState("");
  const file = useRef<HTMLInputElement>(null);
  const [me, setMe] = useState<{ role: string; siteId: string | null } | null>(null);
  const canEdit = me?.role !== "OPERATOR";

  const load = useCallback(async () => {
    const qs = new URLSearchParams(Object.entries(f).filter(([, v]) => v));
    const [e, s, d] = await Promise.all([fetch(`/api/employees?${qs}`), fetch("/api/sites"), fetch("/api/departments")]);
    setEmps(await e.json()); setSites(await s.json()); setDepts(await d.json());
  }, [f]);
  useEffect(() => { load(); }, [load]);
  useEffect(() => { fetch("/api/auth/me").then((r) => r.json()).then((m) => { if (m.role === "SUPER_ADMIN" && !m.mfaEnabled) location.href = "/security"; setMe(m); }); }, []);

  async function call(url: string, method: string, body?: unknown) {
    const r = await fetch(url, { method, headers: { "Content-Type": "application/json" }, body: body ? JSON.stringify(body) : undefined });
    const j = await r.json();
    setErr(r.ok ? "" : j.error ?? "Something went wrong.");
    if (r.ok) load();
    return r.ok;
  }

  async function saveEmployee() {
    const ok = form.id ? await call(`/api/employees/${form.id}`, "PATCH", form) : await call("/api/employees", "POST", form);
    if (ok) setForm(blank);
  }

  async function upload(file: File) {
    const r = await fetch("/api/employees/import", { method: "POST", body: await file.text() });
    const j = await r.json();
    if (!r.ok) setErr(j.error); else { setErr(""); setImp(j); load(); }
  }

  return (
    <div className="dark">
      <main className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-6 space-y-8">
        <header className="flex items-center gap-3">
          <h1 className="text-xl font-semibold mr-auto">Directory</h1>
          <a href="/dashboard" className="text-sm underline">Back to dashboard</a>
        </header>
        {err && <p role="alert" className="rounded-md bg-red-950 text-red-200 px-3 py-2 text-sm">{err}</p>}

        <section className={me?.role === "SUPER_ADMIN" ? "grid md:grid-cols-2 gap-6" : "hidden"}>
          <div className="space-y-2">
            <h2 className="font-semibold">Sites</h2>
            {sites.map((s) => (
              <div key={s.id} className="flex items-center gap-2 text-sm">
                <span className="mr-auto">{s.name} <span className="text-slate-500">sends as "{s.senderId}", {s._count.employees} employees</span></span>
                <button aria-label={`Edit ${s.name}`} onClick={() => { const name = prompt("Site name", s.name); const senderId = name && prompt("Sender ID", s.senderId); if (name && senderId) call(`/api/sites/${s.id}`, "PATCH", { name, senderId }); }}><Pencil size={14} /></button>
                <button aria-label={`Delete ${s.name}`} onClick={() => confirm(`Delete ${s.name}?`) && call(`/api/sites/${s.id}`, "DELETE")}><Trash2 size={14} /></button>
              </div>
            ))}
            <div className="flex gap-2">
              <input className={`${input} w-32`} placeholder="Site name" value={newSite.name} onChange={(e) => setNewSite({ ...newSite, name: e.target.value })} />
              <input className={`${input} flex-1`} placeholder="Sender ID, e.g. OAMPI Davao" value={newSite.senderId} onChange={(e) => setNewSite({ ...newSite, senderId: e.target.value })} />
              <button aria-label="Add site" className={input} onClick={async () => (await call("/api/sites", "POST", newSite)) && setNewSite({ name: "", senderId: "" })}><Plus size={16} /></button>
            </div>
          </div>
          <div className="space-y-2">
            <h2 className="font-semibold">Departments</h2>
            {depts.map((d) => (
              <div key={d.id} className="flex items-center gap-2 text-sm">
                <span className="mr-auto">{d.name} <span className="text-slate-500">{d._count.employees} employees</span></span>
                <button aria-label={`Rename ${d.name}`} onClick={() => { const name = prompt("Department name", d.name); if (name) call(`/api/departments/${d.id}`, "PATCH", { name }); }}><Pencil size={14} /></button>
                <button aria-label={`Delete ${d.name}`} onClick={() => confirm(`Delete ${d.name}?`) && call(`/api/departments/${d.id}`, "DELETE")}><Trash2 size={14} /></button>
              </div>
            ))}
            <div className="flex gap-2">
              <input className={`${input} flex-1`} placeholder="Department name" value={newDept} onChange={(e) => setNewDept(e.target.value)} />
              <button aria-label="Add department" className={input} onClick={async () => (await call("/api/departments", "POST", { name: newDept })) && setNewDept("")}><Plus size={16} /></button>
            </div>
          </div>
        </section>

        <section className="space-y-3">
          <div className="flex flex-wrap items-center gap-2">
            <h2 className="font-semibold mr-auto">Employees ({emps.length})</h2>
            <input className={input} placeholder="Search name or number" value={f.q} onChange={(e) => setF({ ...f, q: e.target.value })} />
            <select className={input} value={f.siteId} onChange={(e) => setF({ ...f, siteId: e.target.value })}>
              <option value="">All sites</option>{sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <select className={input} value={f.deptId} onChange={(e) => setF({ ...f, deptId: e.target.value })}>
              <option value="">All departments</option>{depts.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
            </select>
            <input ref={file} type="file" accept=".csv" hidden onChange={(e) => e.target.files?.[0] && upload(e.target.files[0])} />
            <button className={`${input} ${canEdit ? "flex" : "hidden"} items-center gap-2`} onClick={() => file.current?.click()}><Upload size={14} /> Import CSV</button>
          </div>
          {imp && (
            <div className="rounded-md bg-slate-900 ring-1 ring-slate-800 p-3 text-sm">
              Added {imp.created}, updated {imp.updated}, skipped {imp.errors.length}.
              {imp.errors.length > 0 && <ul className="mt-2 text-amber-400 max-h-40 overflow-y-auto">{imp.errors.map((e) => <li key={e.row}>Line {e.row}: {e.reason}</li>)}</ul>}
            </div>
          )}
          <p className="text-xs text-slate-500">CSV columns: FirstName, LastName, PhoneNumber, Site, Department. Sites must exist first; departments are created automatically.</p>

          <div className={`${canEdit ? "flex" : "hidden"} flex-wrap gap-2 rounded-md bg-slate-900 ring-1 ring-slate-800 p-3`}>
            <input className={`${input} w-32`} placeholder="First name" value={form.firstName} onChange={(e) => setForm({ ...form, firstName: e.target.value })} />
            <input className={`${input} w-32`} placeholder="Last name" value={form.lastName} onChange={(e) => setForm({ ...form, lastName: e.target.value })} />
            <input className={`${input} w-40`} placeholder="09XX XXX XXXX" value={form.phoneNumber} onChange={(e) => setForm({ ...form, phoneNumber: e.target.value })} />
            <select className={input} value={form.siteId} onChange={(e) => setForm({ ...form, siteId: e.target.value })}><option value="">Site</option>{sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select>
            <select className={input} value={form.deptId} onChange={(e) => setForm({ ...form, deptId: e.target.value })}><option value="">Department</option>{depts.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}</select>
            <button className="rounded-md bg-slate-100 text-slate-900 font-medium px-4" onClick={saveEmployee}>{form.id ? "Save changes" : "Add employee"}</button>
            {form.id && <button className={input} onClick={() => setForm(blank)}>Cancel</button>}
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="text-left text-slate-500"><tr><th className="py-2">Name</th><th>Phone</th><th>Site</th><th>Department</th><th /></tr></thead>
              <tbody>
                {emps.map((e) => (
                  <tr key={e.id} className="border-t border-slate-800">
                    <td className="py-2">{e.firstName} {e.lastName}</td><td className="tabular-nums">{e.phoneNumber}</td><td>{e.site.name}</td><td>{e.department.name}</td>
                    <td className={`text-right space-x-3 ${canEdit ? "" : "invisible"}`}>
                      <button aria-label={`Edit ${e.firstName}`} onClick={() => setForm({ id: e.id, firstName: e.firstName, lastName: e.lastName, phoneNumber: e.phoneNumber, siteId: e.siteId, deptId: e.deptId })}><Pencil size={14} /></button>
                      <button aria-label={`Delete ${e.firstName}`} onClick={() => confirm(`Delete ${e.firstName} ${e.lastName}? Their past replies are removed too.`) && call(`/api/employees/${e.id}`, "DELETE")}><Trash2 size={14} /></button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {emps.length === 0 && <p className="py-6 text-sm text-slate-500">No employees match. Add one above or import a CSV.</p>}
          </div>
        </section>
      </main>
    </div>
  );
}
