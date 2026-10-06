# Emergency Response Team (ERT) app

Next.js (App Router) + Tailwind v4 + Prisma/PostgreSQL. Two-way SMS through Globe/m360: broadcasts go out under each site's Sender ID, and employees reply free to shortcode 6780 with SAFE, NOT SAFE or HELP.

**Status:** written but not yet compiled or run. Expect small fixes on first build.

## Quick start
1. `npm install`
2. `cp .env.example .env` and fill it in (generate secrets with `openssl rand -base64 32`).
3. `npx prisma migrate dev --name init`
4. Create the first Super Admin: `SEED_ADMIN_EMAIL=you@company.com SEED_ADMIN_PASSWORD='a long passphrase' npm run seed`
5. `npm run dev`, open `/login`, sign in, and set up two-step sign-in when asked.
6. In **Directory**, add your sites (name + Sender ID, e.g. Davao / "OAMPI Davao"), then import employees by CSV.
7. Point Globe/m360 inbound for shortcode 6780 at `POST https://<host>/api/webhooks/globe-inbound` with header `x-webhook-secret`.

Confirm the m360 outbound URL, auth header, batch limit and inbound payload with Globe. The payloads follow your spec, and `lib/globe.ts` is the one place to adjust.

## Pages
- `/dashboard`: live metrics, reply feed (HELP flashes red), non-responders with call links and "Send again", site tabs, dark mode, double-confirm broadcast.
- `/directory`: sites, departments, employees, filters, CSV import (`FirstName, LastName, PhoneNumber, Site, Department`).
- `/users` (Super Admin): users, password resets, deactivation, two-step resets, activity log.
- `/security`: set up two-step sign-in (authenticator app + recovery codes).

## Roles
| Role | Can do |
|---|---|
| Super Admin | Everything. Must use two-step sign-in. |
| Site Commander | Directory, CSV import and new broadcasts for their own site |
| Operator | View their site; re-send to people who haven't replied |

## Two-step sign-in
- Required for Super Admins (blocked from everything until set up), optional for others.
- Lost phone: use a recovery code, or another Super Admin clicks "Reset two-step" on `/users`. A lone Super Admin who loses both is locked out and needs a database fix, so keep the recovery codes safe, or keep two Super Admins.
- `MFA_ENCRYPTION_KEY` encrypts the stored secrets. Back it up separately. If it's lost, everyone must set MFA up again.

## Notes
- Live updates use in-memory events and rate limits: single instance only. Use Redis or Supabase Realtime for more.
- Not built: Shadcn UI components (plain Tailwind is used), CSV export, delivery receipts from Globe.
- Read `SECURITY.md` before going live.
