# Loading List

Google Apps Script for the bakery loading list spreadsheet (tabs `Owner_List`, `Loading_Board` and `Leftover`).

## How it works

- **Owner_List** is where orders are typed in, one row per order (rows 8–106). A new customer name gets the next Order ID and `Source = Manual`.
- **Loading_Board** has **one row per customer**. Orders with the same customer name are added together. Names match ignoring capital letters and extra spaces, so `Uche` and `uche ` are the same customer. The board is rebuilt from Owner_List whenever Owner_List changes, including when rows are inserted, deleted or sorted.
- Loaders set the **Status** on Loading_Board. The status, the time it changed and the row colour are copied to every Owner_List row for that customer.
- A customer marked **Leftover** gets an Order / Taken / Leftover block on the **Leftover** tab. The Order row follows the board. Type what was taken into the Taken row, and Leftover = Order − Taken. Tick the checkbox in column P to stop a leftover carrying into the next shift. Negative amounts are allowed everywhere.

### 🍞 Bakery Tools menu

| Item | What it does |
| --- | --- |
| Sync All Statuses & Colors | End of shift: every customer still `Not started` becomes `Leftover`, and the Leftover tab is brought up to date. |
| 🔄 Auto-Correct Missing Orders | Rebuilds the board from Owner_List, rewrites the totals rows and fixes the Leftover tab. Safe to run any time. |
| ⚠️ Reset Board & Leftovers | Sets every status back to `Not started`, erases timestamps and empties the Leftover tab. Orders are kept. |
| Generate Next Shift Sheet | Copies the file into `Loading List/<year>/<month>` in Drive for the next shift (Sunday Morning is skipped), fills its Owner_List with the unticked leftovers, installs its triggers and emails the link to the list in Owner_List `H2`. Leftover `N2`/`O2` override the next date and shift. |
| Install Triggers For This File | Sets up the edit and change triggers for the open file. |

## Installing or updating

1. In the spreadsheet, open **Extensions → Apps Script** and replace the code with [`Code.gs`](Code.gs). Save.
2. **Remove the old triggers first.** Each generated file was wired up by the script of the file it was copied from, so the old code can still be running on it. Go to <https://script.google.com/home/triggers> and delete the `installedOnEdit` triggers that point to this spreadsheet.
3. Reload the spreadsheet and choose **🍞 Bakery Tools → Install Triggers For This File**. Approve the permissions.
4. Choose **🍞 Bakery Tools → 🔄 Auto-Correct Missing Orders** once so the board merges the existing orders.

Files you generate from this one get the new code's triggers automatically.

## Tests

`node test/run.js` runs the script against an in-memory stand-in for Google Sheets (see `test/mock.js`). It needs Node 18 or newer.
