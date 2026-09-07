# Loading List — Apps Script

Container-bound Apps Script project for the bakery Loading List spreadsheet
(`Owner_List`, `Loading_Board`, `Leftover` tabs).

## Files

- `Code.gs` — existing sheet automation: syncing `Owner_List` ↔ `Loading_Board`,
  leftovers, status colors, the `🍞 Bakery Tools` menu, and shift rollover.
- `WebApp.gs` — server-side code for the New Order web app (`doGet`, `getFormMeta`,
  `submitOrder`).
- `OrderForm.html` — the web app's form UI.
- `appsscript.json` — project manifest (timezone, web app defaults).

## New Order web app

Lets staff add a new order to `Owner_List` (the first tab) from a phone or
browser, without opening the spreadsheet. On submit it:

1. Finds the next blank row in `Owner_List` (rows 8–106).
2. Writes the customer name and bread quantities.
3. Sets `Source` to `Manual` and assigns the next Order ID via the existing
   `generateOrderId()`.
4. Pushes the row onto `Loading_Board` via the existing `syncSingleRowToBoard()`.

This mirrors exactly what happens when someone types a new row into
`Owner_List` by hand, so `installedOnEdit`, status colors, and the Leftover
workflow all keep working unchanged.

The customer dropdown and the bread-quantity columns/groups are read live
from the sheet (`B8`'s data validation, and the `TASTY`/`TOSBI`/`CHAMPION`
header rows), so editing those on the sheet updates the form automatically —
no code changes needed.

### Deploying

This repo has no `.clasp.json` (no script ID is checked in). To deploy:

1. Install [`clasp`](https://github.com/google/clasp) and run `clasp login`.
2. From the Apps Script editor already bound to the spreadsheet, either:
   - `clasp clone <scriptId>` into this folder and copy these files in, or
   - open **Extensions → Apps Script** on the spreadsheet and paste in
     `Code.gs`, `WebApp.gs`, `OrderForm.html`, and `appsscript.json` by hand
     (create matching files/types in the editor).
3. **Deploy → New deployment → type: Web app.**
   - Execute as: **Me** (so staff without edit access to the sheet can still
     submit orders through your permissions) — this is set as the default in
     `appsscript.json`, but confirm it in the deploy dialog.
   - Who has access: pick based on your org (**Anyone** for anyone with a
     Google account, **Anyone within [domain]** for a Workspace domain-only
     tool). Adjust in the deploy dialog regardless of the manifest default.
4. Copy the resulting web app URL and share it with staff.

Re-deploying after edits requires **Manage deployments → edit → new version**
(a plain save does not update a live web app URL).
