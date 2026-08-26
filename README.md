# Church Event Registration MVP

Google Forms + Google Sheets + Apps Script registration, QR check-in, and meal redemption. Each attendee receives one QR code. Event and meal names are configuration data rather than code.

## Included workflow

1. Organizer configures an event and its meal choices.
2. Attendee submits one Google Form per person.
3. The system creates a registration, emails a QR receipt, and sends an SMS receipt link through Twilio.
4. Volunteers scan the same QR in check-in or meal-pickup mode.
5. Meal pickup succeeds only when a meal was selected, the organizer marked it ordered, and it has not already been redeemed.

## Initial deployment

1. Create a new standalone Apps Script project at `script.google.com`.
2. Add every `.gs` and `.html` file in this folder. `ScannerPage.html` intentionally has a different base name from `Scanner.gs`, because the Apps Script browser editor requires unique base names. Replace the generated manifest with `appsscript.json` after enabling **Show appsscript.json manifest file** in Project Settings. Alternatively, use `clasp` and the example configuration.
3. Run `setupApplication()` once and approve the requested Google permissions. Its execution result contains the spreadsheet and form URLs. Keep the form closed for now.
4. In the generated **Meals** sheet, replace the sample foods. Keep `No meal` if it should be offered, then run `syncMealChoicesToForm()`.
5. In **Settings**, enter the event name, ID, date, location, and organizer email.
6. Run `setOperatorPins('123456', '654321')` with new private PINs. Do not use these examples.
7. Deploy as **Web app**: execute as yourself and allow access to anyone with the URL. Copy the `/exec` deployment URL.
8. Run `setWebAppUrl('YOUR_EXEC_URL')`, then `activateRegistration()`.
9. Configure Twilio with `setTwilioCredentials(accountSid, authToken, fromNumber)`. Until then, email works and SMS is logged as skipped.

The public registration URL is the form's published URL. Scanner URLs are:

- Check-in: `WEB_APP_URL?view=scanner&mode=checkin`
- Meal pickup: `WEB_APP_URL?view=scanner&mode=meal`

## Organizer operation

- Review **Food Summary** for requested counts.
- After placing the food order, run `finalizeMealOrder()`. This marks every active non-`No meal` selection as ordered.
- A registration can be changed individually in **Registrations** if an exception is required.
- Keep the two operator PINs private and separate.
- Review **Scan Log** for successful and rejected scans.

## Test

Run `npm test` from this directory. These tests validate phone normalization, QR parsing, check-in duplication, meal-order enforcement, and duplicate meal redemption.

## Important MVP notes

- Twilio charges for messages and requires appropriate account/number configuration. The form includes explicit transactional SMS consent.
- QR images use QuickChart in this MVP; the QR contains only a random registration token, not the attendee's name, email, phone, or food choice. A production version can vendor a QR encoder to remove this external image dependency.
- Apps Script and email have daily quotas. Confirm the expected attendance against the quotas of the deploying Google account before a large event.
- Deploying for anonymous access makes the receipt and scanner pages reachable by URL. Scanner mutations still require a mode-specific PIN. For higher-security or larger events, migrate operator authentication and data storage to a managed backend.
- Treat the spreadsheet as sensitive personal data and share it only with authorized organizers.

## Suggested phone acceptance test

1. Submit three registrations: regular meal, special/dietary note, and `No meal`.
2. Confirm email and SMS receipt links on iPhone and Android.
3. Scan every QR at check-in; rescan one and verify the duplicate warning.
4. Attempt lunch scans before finalizing the meal order; verify they are rejected.
5. Run `finalizeMealOrder()` and retry; valid meals should pass.
6. Rescan a redeemed meal; verify it reports the earlier redemption.
7. Scan the `No meal` registration; verify it is rejected.
8. Turn off Wi-Fi temporarily and verify the desk understands that this MVP requires connectivity.
