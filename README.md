# CRES Portal — Internal View (test build)

Copy of `CRES-TRIAL-PORTAL` with three additions. Everything else (days, categories,
logistics, versions, calculation, client sheet layout) works exactly as before.

## What's new

1. **Vendor Amount column** right after Amount, in every item table and in Logistics.
   - Typed directly (accepts sums like `15000*3`). Everything else in the row is shared.
   - Summary card shows a Vendor box: vendor subtotal, vendor grand total (same mgmt fee,
     hotel and tax as the client, but no management fee), and margin.
2. **Two sheets on submit**, identical layout:
   - `Estimate - <event> - <client>` titled **CRES — CLIENT COSTING**, Amount column = Amount.
   - `Vendor Costing - <event> - <client>` (separate file, same folder) titled
     **CRES — VENDOR COSTING**, Amount column = Vendor Amount. Same V1/V2 tabs.
   - A row that has only a Vendor Amount (nothing else) appears only on the vendor sheet.
   - Delete version / delete event also removes the vendor tab / file.
3. **How to fill** guidance per library item (Library column D). Tap the small **i** next to an
   item, or tick *Show "How to fill"*.

Also fixed:
- Clicking or tabbing through the Amount box no longer freezes the auto amount (live portal has this bug).
- Very large events no longer hit the 50,000-character cell limit when saving (hidden `_Data_`
  snapshot is split across cells; old events still load).

## Test setup (does not touch the live portal)

1. Open the live portal spreadsheet → **File → Make a copy**. The Apps Script is copied with it.
2. In the copy, **clear Index rows 2 onwards** (keep the header row). Those rows point to the
   live event files — editing them from the test portal would change live estimates.
3. In the copy: **Extensions → Apps Script** → replace all of `Code.gs` with the `Code.gs` from
   this repo. For testing, change line 1 to `var FOLDER_NAME = "CRES Estimates TEST";` so test
   files land in their own folder.
4. Run **`setup`** once (adds the Index internal columns and Library "How to fill" header),
   then run **`seedHowToFill`** (fills guidance for every item; it never overwrites text you
   have already typed in column D).
5. **Deploy → New deployment → Web app**, Execute as: *Me*, Who has access: *Anyone* →
   copy the `/exec` URL.
6. In `index.html`, set `const SCRIPT_URL = "...";` to that URL.
7. Upload `index.html`, `manifest.json`, `icon-192.png`, `icon-512.png` to the repo root and turn on
   **Settings → Pages → Deploy from branch → main / root**.
   The page will be at `https://mis0303.github.io/CRES-Portal-Internal/`.

## Going live later

1. Paste this `Code.gs` into the live portal's Apps Script (keep `FOLDER_NAME = "CRES Estimates"`),
   run `setup` and `seedHowToFill`.
2. **Deploy → Manage deployments → edit the existing deployment → New version**, so the `/exec`
   URL stays the same.
3. Copy this `index.html` into `CRES-TRIAL-PORTAL` with the live `SCRIPT_URL`.

Old events open normally; their Vendor Amount column starts empty and the vendor file is created
the first time they are saved.

## Editing guidance

Edit text directly in Library column D. To regenerate everything from the list in `Code.gs`,
run `seedHowToFillOverwrite` (this replaces your edits).

## Note on access

The portal has no login, and the web app is open to anyone with the `/exec` URL. That URL is in
the page source, so vendor numbers can be read by anyone who has the portal link.
Add a PIN or Google sign-in before sharing widely.
