# RCCC Event Registration — project guide

Multi-event registration, QR ticketing, and check-in system for RCCC (Rutgers
Community Christian Church). Zero-infrastructure by design: Google Apps Script
backend + Google Sheets storage + Google Forms registration + GitHub Pages
front-ends + one Cloudflare Worker (Apple pass signing).

**This repo is public. Never commit PINs, tokens, keys, or certificates.**
Secrets live outside the repo in `~/Documents/RCCC/wallet-secrets/` and in
Apps Script Script Properties. Operator/admin PINs live in the admin console
and each event's Stations sheet.

## Architecture

- **Master spreadsheet** (Apps Script property `SPREADSHEET_ID`): `Settings`,
  `Users` (global identity: verified Google email ↔ one `user_id`), `Events`
  (control surface), `Meal Template`.
- **Per-event spreadsheet + Google Form**, provisioned by `applyEventChanges()`
  or the admin console: `Registrations`, `Meals`, `Scan Log`, `Food Summary`
  (has `remaining` col), `Stations` (station_id, station_name, type
  checkin|meal, pin, active, location, info_url), `Check-ins` (per-station
  duplicate detection).
- **PIN-as-router**: one universal scanner page; the station PIN alone resolves
  to (event, station, mode). Closing an event retires its PINs. Tickets show
  identity only (guest + registration + QR) — entitlements (meal, sessions)
  resolve at scan time at the station.
- **Self-service updates**: resubmitting the form with the same Google account
  + same attendee name updates the registration in place (same QR token; a
  changed meal clears `meal_ordered`).

## Live pieces

| Piece | Where |
|---|---|
| Apps Script backend (doGet/doPost) | script project `1ZuuX-q5NtK-EvpOYa3nRVenRnowsEYOlIjBXwD2Uj7RiLYAz7yVyLuMq`, deployed under jiruiryanhan@gmail.com |
| Web app /exec (admin `?view=admin`, receipt `?view=receipt`, HTTP API via doPost) | deployment `AKfycbytJhW6z4YiRgiGV6YnuZ9uOoJBhl9eRiaDmYd1c07PAjC053i1gneCyOOX68Cg5MKFUQ` |
| Scanner + landing page (GitHub Pages, `docs/`) | https://jrhan-p.github.io/EventRegistration/ , `scanner.html` |
| Apple pass signer (Cloudflare Worker, `worker/`) | https://rccc-pass.rccc-events.workers.dev (church CF account `9cc803d07181597767e37bc0cc95a6b4`) |
| Google Wallet | Issuer `3388000000023178444`, class `rccc-events` (demo mode; publishing pending) |
| Apple Wallet | Team `48UG39GUGA`, Pass Type ID `pass.org.rccc.events`, cert exp 2027-09 |

## Deployment rules (hard-won — do not relearn these)

1. **Apps Script web deployments happen ONLY in the editor UI** (Deploy →
   Manage deployments → pencil → New version). `clasp create-deployment` /
   `clasp update-deployment` silently strip the web-app entry point → /exec
   404s. `clasp push` alone is fine and is all trigger-side code needs
   (triggers run HEAD; doGet/doPost + HTML run the deployed version).
2. **Never use `google.script.run`** in served pages — it fails silently when
   browsers block third-party cookies. All pages use plain `fetch` POSTs to
   the doPost API (form-encoded → no CORS preflight; the 302-to-
   googleusercontent hop carries ACAO:*).
3. **Never name POST params `sid` and `token` together** — Google's /exec
   front end hijacks such requests before the script runs.
4. **Camera (getUserMedia) does not work inside Apps Script's iframe wrapper**
   on iOS WebKit. That's why the scanner lives on GitHub Pages, top-level.
5. Mobile viewport on Apps-Script-served pages requires
   `HtmlOutput.addMetaTag('viewport', …)` (`page_()` in Web.gs).
6. **Google Wallet classes created in the console default to DRAFT and cannot
   issue objects** — set Status to UNDER_REVIEW (auto-APPROVED in demo).
7. **curl against /exec**: use `curl -L --data …` WITHOUT `-X POST` (forcing
   POST breaks the 302 GET hop and returns HTML instead of JSON).
8. **Shared machine**: parallel Claude sessions run other orgs' accounts here.
   Never `gh auth switch`, never `wrangler login/logout`, never bare
   `wrangler deploy`. Deploy the worker with
   `CLOUDFLARE_API_TOKEN=$(cat ~/Documents/RCCC/wallet-secrets/cf-api-token.txt) CLOUDFLARE_ACCOUNT_ID=9cc803d07181597767e37bc0cc95a6b4 npx wrangler deploy`
   (run inside `worker/`; local wrangler 3.x — machine Node is v20). For gh:
   `GH_TOKEN=$(gh auth token --user jrhan-p) gh …`. Git push needs nothing
   (remotes are URL-pinned per account with scoped credential helpers).
9. GitHub: `origin` = jrhan-p/EventRegistration (church account),
   `upstream` = davidchaozhang/EventRegistration (read-only). Commit identity
   is repo-local (jrhan-p noreply). Never touch global git identity.

## Routine workflows

- **Backend change**: edit `.gs` → `npm test` → `npx clasp push -f` → if it
  affects doGet/doPost/HTML: editor-UI redeploy (New version). Trigger-only
  changes need no redeploy.
- **Scanner/landing change**: edit `docs/` → commit → `git push origin main`
  (Pages redeploys in ~1 min).
- **Worker change**: edit `worker/src/` → env-token `npx wrangler deploy`
  (rule 8). Secrets via `npx wrangler secret bulk` from the secrets dir.
- **Scan perf**: platform floor is ~1s; scans run ~2s warm. Lock only wraps
  dup-check + writes; token→row, header, PIN→station lookups are cached in
  CacheService (closing an event leaves its PINs valid ≤5 min).

## Testing

`npm test` runs `tests/core.test.js` (pure logic in Core.gs). End-to-end:
register via the event form (needs Google sign-in — a human), check the email,
scan with a station PIN from the admin console. `fn=passData`, `fn=scan`,
`fn=resolvePin` etc. are curl-testable against /exec.

## Roadmap / parked

- Google Wallet publishing access: payments profile (owner's legal info) →
  Request publishing access → days of review; removes [TEST ONLY] and the
  tester allowlist.
- Church domain: CNAME to GitHub Pages for the scanner/landing; later a custom
  domain for the Worker. Zero code change except `SCANNER_BASE` (Setup.gs).
- Google Workspace via Google for Nonprofits (free): church noreply@ sender,
  1,500 emails/day (consumer Gmail caps ~100/day).
- Twilio SMS: wired end-to-end, credentials configurable in the admin console;
  parked to stay at $0. Production US texting needs upgrade + A2P 10DLC.
- Pass artwork polish (Apple strip image, Google hero image).
- Sub-second scanning (move the scan endpoint into the Worker) if a large
  event demands it.
