# Church Event Registration MVP

Google Forms + Google Sheets + Apps Script registration, QR check-in, and meal redemption. Each attendee receives one QR code. The system is multi-event: every event gets its own spreadsheet, its own customizable Google Form, and its own scanner links, all provisioned from one master control sheet.

Registration requires signing in with a Google account. Google attaches the verified account email to every submission, and the system maps each verified email to exactly one `user_id` in the master **Users** sheet (strictly one-to-one; no heuristic merging). One user may submit several registrations — for example, one per family member — and every registration row carries its owner's `user_id`. Users are global: the same person keeps the same `user_id` across all events.

## Architecture

- **Master spreadsheet** (created by `setupApplication()`): `Settings` (organizer email, web app URL), `Users` (global identity), `Events` (the control surface), `Meal Template` (default menu for new events).
- **Per-event spreadsheet** (created per Events row): `Registrations`, `Meals`, `Scan Log`, `Food Summary`, plus the linked form responses.
- **Per-event Google Form**: generated from the meal template, then freely customizable in the Forms editor. Keep the titles of the fields the code reads (`Full Name`, `Mobile Phone`, `Food Selection`, `Dietary Notes`, `SMS Consent`); extra questions are allowed and simply stay in the form responses tab.

## Organizer workflow

1. Open the master **Events** sheet and add a row: `event_name`, `event_date`, `location` (leave `status` empty).
2. Run `applyEventChanges()` in the Apps Script editor. The row fills in with the public registration URL, both scanner URLs, and links to the event's spreadsheet and form.
3. Customize the event's `Meals` sheet if needed, then run `syncAllMealChoices()`; or edit the form directly.
4. Share the registration URL. Attendees sign in with Google, register, and receive a QR receipt by email (and SMS when Twilio is configured).
5. Review the event's **Food Summary**, place the food order, then type `finalize` in the event's `action` column and run `applyEventChanges()`.
6. Volunteers open the scanner links (check-in and meal modes have separate PINs). Meal pickup succeeds only when a meal was selected, marked ordered, and not yet redeemed.
7. To stop registrations, type `close` in the `action` column and run `applyEventChanges()` (`reopen` reverses it). To discard an event entirely, delete its spreadsheet and form in Drive and its row in Events.

`cleanupLegacyData()` removes the tabs, test users, and form left over from the original single-event deployment.

## Initial deployment

1. Create a standalone Apps Script project (or use `clasp` with this folder) and push every `.gs` and `.html` file plus `appsscript.json`.
2. Run `setupApplication()` and approve permissions.
3. Run `setOperatorPins('...', '...')` with two private 4–10 digit PINs.
4. Deploy as **Web app** (execute as you; access: anyone) **from the editor UI** — see the warning below. Visit the `/exec` URL once; the `WEB_APP_URL` setting self-heals to match it.
5. Add the first event row and run `applyEventChanges()`.
6. Optional SMS: run `setTwilioCredentials(accountSid, authToken, fromNumber)`.

> **Deployment warning:** the Apps Script API (and therefore `clasp deploy` / `clasp update-deployment`) silently strips the web-app entry point — the `/exec` URL then returns 404. Always create and update web-app deployments in the editor UI (Deploy → Manage deployments → edit → New version). `clasp push` alone is fine and is all that trigger-side code needs.

## Test

Run `npm test` from this directory. These tests validate phone normalization, QR parsing, check-in duplication, meal-order enforcement, and duplicate meal redemption.

## Important MVP notes

- Twilio charges for messages and requires appropriate account/number configuration. The form includes explicit transactional SMS consent.
- QR images use QuickChart in this MVP; the QR contains only a random registration token, not the attendee's name, email, phone, or food choice.
- Apps Script and email have daily quotas (consumer Gmail ≈ 100 recipients/day). Confirm expected attendance against the deploying account's quotas before a large event.
- Scanner mutations require a mode-specific PIN. Treat the spreadsheets as sensitive personal data and share them only with authorized organizers.
