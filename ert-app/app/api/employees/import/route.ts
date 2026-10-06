import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db";
import { parseCsv } from "@/lib/csv";
import { normalizePhone } from "@/lib/globe";
import { clientIp, guard, requireUser, siteScope } from "@/lib/auth";
import { audit } from "@/lib/audit";

// Body: raw CSV text. Columns: FirstName, LastName, PhoneNumber, Site, Department.
// Sites must already exist (they carry a Sender ID). Departments are created on the fly.
// Existing phone numbers are updated, so re-importing a corrected list is safe.
export const POST = guard(async (req: NextRequest) => {
  const sess = await requireUser("SUPER_ADMIN", "SITE_COMMANDER");
  const scope = siteScope(sess);
  const text = await req.text();
  if (text.length > 2_000_000) return NextResponse.json({ error: "The file is over 2 MB. Split it into smaller files." }, { status: 413 });
  const rows = parseCsv(text);
  if (!rows.length) return NextResponse.json({ error: "The file has no rows." }, { status: 400 });
  if (rows.length > 3000) return NextResponse.json({ error: "Import up to 3,000 rows at a time." }, { status: 400 });

  const sites = new Map((await prisma.site.findMany({ where: scope ? { id: scope } : {} })).map((s) => [s.name.toLowerCase(), s.id]));
  const depts = new Map((await prisma.department.findMany()).map((d) => [d.name.toLowerCase(), d.id]));
  let created = 0, updated = 0;
  const errors: { row: number; reason: string }[] = [];

  for (const [i, r] of rows.entries()) {
    const line = i + 2; // header is line 1
    const phone = normalizePhone(r.phonenumber ?? "");
    const siteId = sites.get((r.site ?? "").toLowerCase());
    if (!r.firstname || !r.lastname) { errors.push({ row: line, reason: "Missing name" }); continue; }
    if (!phone) { errors.push({ row: line, reason: `Invalid phone "${r.phonenumber}"` }); continue; }
    if (!siteId) { errors.push({ row: line, reason: `Unknown site "${r.site}" (or not your site). Sites are added by a Super Admin.` }); continue; }
    if (!r.department) { errors.push({ row: line, reason: "Missing department" }); continue; }

    let deptId = depts.get(r.department.toLowerCase());
    if (!deptId) {
      deptId = (await prisma.department.create({ data: { name: r.department } })).id;
      depts.set(r.department.toLowerCase(), deptId);
    }
    const data = { firstName: r.firstname, lastName: r.lastname, siteId, deptId };
    const existing = await prisma.employee.findUnique({ where: { phoneNumber: phone } });
    if (existing) { await prisma.employee.update({ where: { id: existing.id }, data }); updated++; }
    else { await prisma.employee.create({ data: { ...data, phoneNumber: phone } }); created++; }
  }
  await audit("employees.import", { userId: sess.id, email: sess.email, ip: clientIp(req), detail: { created, updated, skipped: errors.length } });
  return NextResponse.json({ created, updated, errors: errors.slice(0, 200) });
});