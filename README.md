# Parcel Ledger

Follow every parcel by its **AWB**, from the packing table to the bank: dispatch scan → Flipkart / Meesho order → delivery or return → settlement received → claim.

It works on top of your two existing apps and changes neither of them:

| App | Where its data lives | What the ledger reads |
|---|---|---|
| Dispatch Scan Register | Google Sheet (Apps Script) | Dispatch scans, cancelled parcels, packet codes |
| Godown Stock | Firebase | Scan-outs with AWB, returns, stock (read-only) |
| **Parcel Ledger** (this app) | Same Google Sheet, new **Ledger** and **Rates** tabs | Orders, status, sale price, settlement, returns, claims |

Flipkart orders, status, sale price and SKU come in **by themselves every hour** through the Flipkart API. Flipkart has no API for payments, so you import the **Settlement report**. Meesho has no seller API, so you import its order, return and payment reports.

---

## Part 1 – Update the Dispatch Google Sheet script (one time, 5 minutes)

The new `google-sheet-script/Code.gs` is your current dispatch script **plus** the ledger. The Dispatch Scan Register keeps working exactly as before.

1. Open the **Dispatch Register** Google Sheet → **Extensions → Apps Script**.
2. Select everything in `Code.gs`, delete it, paste the whole new `google-sheet-script/Code.gs`, click **Save**.
3. Pick **setup** in the function list, click **Run**. It adds the **Ledger** and **Rates** tabs.
4. **Deploy → Manage deployments → ✏️ → Version: New version → Deploy.** The `/exec` link stays the same, so no phone needs changing.

## Part 2 – Switch on Flipkart (one time)

1. Flipkart Seller Hub → **Manage Profile → Developer Access → Self Access**. Copy the **Application ID** and **Application Secret**.
2. In Apps Script: **Project Settings** (gear, left side) → **Script Properties → Add script property**:
   - `FK_APP_ID` = Application ID
   - `FK_APP_SECRET` = Application Secret
   Never paste the keys into the code, into chat, or into GitHub.
3. Back in the editor, pick **checkFlipkart** and **Run**. Open **Execution log**. It should say `SUCCESS`. If it says the keys are swapped or rejected, it tells you what to fix.
4. Run **setup** once more. It pulls the last 30 days of Flipkart orders into **Ledger** and sets an hourly sync.

Running setup with keys saved also switches on the existing 15-minute pre-dispatch order sync (the **Orders** tab). The dispatch app ignores it while `PORTALS = false`.

## Part 3 – Put the ledger on GitHub Pages (one time)

1. On github.com create a new **public** repository named `parcel-ledger`.
2. **Add file → Upload files** and drop in: `index.html`, `app.js`, `style.css`, `manifest.json`, `icon.svg`, `README.md`. Leave out `google-sheet-script` (it is already in the Sheet, and it is fine to upload it too: it contains no keys).
3. **Settings → Pages → Branch: main, folder: / (root) → Save.** After a minute the app is at
   `https://ekanshj2003.github.io/parcel-ledger/`

## Part 4 – Let the ledger read Godown Stock (one time)

1. Firebase console → project **godown-stock-9fcaa** → **Authentication → Settings → Authorized domains**. Make sure `ekanshj2003.github.io` is in the list (add it if not).
2. Your Firestore rules must let a logged-in staff user read `products` and `movements` (they already do if Godown Stock works for staff).

## Part 5 – First login

1. Open the ledger link. Tap **Connect to the Google Sheet**, paste the same `/exec` link the dispatch phones use, **Save and test**.
2. Log in with a name and PIN from the Sheet's **Users** tab. Admins can change rates, sync Flipkart and delete ledger rows.
3. **Connections** tab → log in with a Godown Stock user ID and password. Paste the Godown Stock app link so its button shows at the top.
4. **Rates** tab: SKUs from your parcels show here with *No rate*. Tap each and enter what Flipkart / Meesho pays you per packet after fees. This makes the expected settlement and the short-payment check work.

## Daily / weekly use

- **Ledger**: every parcel with its status, sale price, expected and received money. Coloured boxes at the top filter to what needs attention: unpaid after 15 days, short-paid, returns on the way, stuck in transit, shipped but never scanned, wrong product.
- Click a parcel to mark a return as received (good, damaged, wrong item, never came back), fix a settlement, or **raise a claim**.
- **Claims**: parcels that should be claimed (short-paid, lost, bad or missing returns, unpaid after 45 days), and every claim with its ID, status and money recovered.
- **Reconcile**: four checks for the dates at the top:
  1. Godown out = dispatched (AWB by AWB)
  2. Dispatched = billed (scans vs portal orders, wrong SKU)
  3. Billed = money received
  4. Returns = back in the godown
- **Import** (weekly):
  - Flipkart **Settlement report** → money received, fees, UTR
  - Meesho **orders report** first (links AWB to Sub Order No), then **returns**, then **payments**

Everything is saved in the Google Sheet's **Ledger** tab, so you can also open, filter or share it as a sheet.

## Problems

| What you see | What to do |
|---|---|
| "The Sheet still has the old script" | Part 1: paste the new Code.gs, run setup, deploy a **new version**. |
| Connections says "Flipkart keys are not saved" | Part 2. |
| Last sync failed: HTTP 401 / keys rejected | Run **checkFlipkart** and follow what the log says. |
| "rows had only an order ID that is not linked" | Import the orders report (with AWB) first, then the payment file again. |
| Godown Stock: "User ID or password is wrong" | Use the same login as the Godown Stock app. |
| Godown Stock login fails with "unauthorized domain" | Part 4, step 1. |
