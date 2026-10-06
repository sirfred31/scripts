# Security review

## Fixed in this pass
| Risk | Fix |
|---|---|
| Missing/short `AUTH_SECRET` could yield a forgeable session key | Auth fails closed unless the secret is 32+ chars (app and middleware) |
| Sessions stayed valid after password reset, deactivation or role change | Per-user `sessionVersion` in the token; any such change signs the user out everywhere |
| Login timing revealed which emails exist; lockout let strangers lock real users out | bcrypt always runs; limits keyed by IP+email plus a per-IP cap |
| No record of who did what | `AuditLog`: logins, broadcasts, user changes, deletes, imports (no passwords or phone numbers) |
| A stolen account could spam the whole company | Max 5 broadcasts per user per 10 min, 320-char message cap, duplicate-send block |
| Webhook secret compared with `!==` and open if unset | Constant-time compare, fails closed, optional IP allow-list, shortcode check, body cap, server-side timestamps |
| Upstream Globe error text returned to the browser | Only counts returned; details go to server logs. Failed sends are now shown to the sender |
| Cross-site writes | Origin / Sec-Fetch-Site check on all API writes, on top of SameSite=Lax cookies |
| Clickjacking of the broadcast button, XSS blast radius | CSP, `X-Frame-Options: DENY`, nosniff, HSTS, no-referrer (`next.config.mjs`) |
| Scoped users could read other sites' campaigns by id | Campaign lookups limited to all-sites or the user's own site |
| Oversized CSV / unbounded inputs, unhandled errors leaking stack traces | 2 MB / 3,000-row cap, field limits, generic 500 responses |
| Last Super Admin could be removed | Blocked, and a user can't demote or deactivate themselves |

Checked and fine: no raw SQL (Prisma parameterizes), React escapes reply text, `tel:` links use normalized numbers, cookies are httpOnly and Secure in production.

## Added: two-step sign-in (TOTP)
- Authenticator-app codes (RFC 6238) after the password. Required for Super Admins, who are blocked from the app until it's on.
- Secrets are AES-256-GCM encrypted with a separate `MFA_ENCRYPTION_KEY`. Each code works once (replay and race protected). Five wrong codes lock that account's second step for 15 minutes.
- 8 single-use recovery codes, stored as bcrypt hashes. Enabling, disabling or resetting MFA signs the user out everywhere and is logged.
- Phishing-resistant options (passkeys / WebAuthn) would be a further step up.

## Still your call (not code)
- **Rate limits and lockout are in memory.** Use Redis if you run more than one instance. `x-forwarded-for` is only trustworthy behind your own proxy.
- **Rotate secrets** (`AUTH_SECRET`, Globe token and webhook secret) if a laptop, log or chat ever exposed them.
- **Database:** least-privilege DB user, encrypted backups, TLS to Postgres, restrict network access.
- **Employee phone numbers are personal data.** Check Philippine Data Privacy Act obligations (consent notice, retention, access control) with your DPO.
- **Dependencies:** commit the lockfile, run `npm audit` in CI, enable Dependabot.
- **Test the emergency path** with a drill, and have a manual fallback if SMS or the app is down.
- **Independent review** (penetration test) before real use. This review was by reading the code, not running it.
