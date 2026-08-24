# ShedSuite → RTO Pro Web App V7

V7 adds immediate per-contract Discard on the review screen and normalizes 11-digit US phone numbers beginning with country code 1 before applying the existing phone comparison rule. It also retains the V6.3.5 theme fixes and expanded hidden Cyberpunk palette.

V7 is built on the V3 workflow of the local browser replacement for `Contract_Import.xlsm` + `Contract_Import.py` from **Shedsuite_Import V1.1.7.6**.

## V7 background Delivery Certificates

- Review/edit opens after the fast ShedSuite CSV/direct-PDF work; Delivery Certificate Chromium no longer blocks the page.
- Chromium runs **headless in a background worker** and reuses login contexts during the job.
- Review page shows live queued/downloading/ready/missing progress without rerendering or wiping unsaved edits.
- Missing certificates can be retried; queued/missing certificates can be skipped.
- Discarding a contract causes the background worker to skip it and removes any packet created after discard.
- Certificate completion merges only server-owned PDF metadata into the latest saved row, preventing background work from overwriting user edits.
- Saving edits also preserves newer background certificate metadata if the browser loaded an older copy.


## V3 review workflow

The overall card UI remains the same, but each contract is now split into focused tabs:

- **Mapping** — Store, Dealer, Dealer ID, Account, Model #, Contract #, Zone, quick tax status.
- **Inventory** — Category and Description are directly next to each other, followed by model/serial/condition/brand/vendor/cost/retail/rate/invoice fields.
- **Customer** — customer identity/contact fields, Mailing Address and Delivery Address side-by-side, references, phone selection rule, and editable `COMMENTS`.
- **Tax & Address** — delivery/tax address, automatic ZipTax result, RTO Tax Zone dropdown/manual override, tax rate, county and zone.
- **PDF / Agent** — extracted PDF Salesman/Agent, RTO Firebird Agent, invoice number, extracted GPS coordinates, Directions and a link to the combined PDF.

The **All RTO fields** and **Source row** buttons remain available on every card.

## Phone rule

V3 implements the requested phone behavior:

1. Primary Phone always maps to the RTO Primary/Cell field.
2. If Secondary is empty, use Ref 1 as the RTO Other phone.
3. If Secondary matches Primary, use Ref 1.
4. If Secondary matches Ref 2, use Ref 2.
5. If Secondary matches Ref 1, use Ref 1.
6. Otherwise use Secondary.

The Customer tab shows the selected source (`Ref 1`, `Ref 2`, or `Secondary`) and allows a manual override. **Reapply auto rule** returns it to automatic logic.

When Ref 1 is selected, `COMMENTS` becomes `OTHER IS <RELATIONSHIP> <NAME>`. Ref 2 works the same way. An unmatched Secondary produces `OTHER IS SECONDARY PHONE`. The comment box remains editable.

## ZipTax

ZipTax now runs automatically from the **Delivery Address** during the initial import. V3 preserves the original desktop program's local ZipTax behavior and allows an environment-key override through `.env` if needed.

The app uses ZipTax county/city/incorporation/rate information to choose the best matching active RTO Pro Tax Zone from Firebird. The result is shown in the Tax tab, and the RTO Tax Zone remains a dropdown so you can change it manually.

If you edit the delivery address, click **Detect tax zone with ZipTax** to run the lookup again.

## PDF Salesman / Agent extraction

V2 usually downloaded only the signed RTO contract. In the supplied documents, the salesperson is on the **Invoice PDF** under `Agent:`. V3 therefore always includes the Invoice PDF in the collection used for extraction.

The extractor now accepts multiple ShedSuite layouts including `Agent:`, `Agent`, `Salesman:` and same-line values. The extracted value fills RTO `SALESMAN`. `AGENT1` remains the separate RTO Pro/Firebird agent associated with the selected Store, matching the workbook design.

GPS coordinates from the delivery/happy-sheet PDF are stored separately and displayed in the **PDF / Agent** tab. They are also included in Directions when found.

## First page / Firebird settings

The first page now has two setup tabs:

- **Import** — CSV and PDF options.
- **Firebird settings** — DSN/user/password.

The Firebird connection fields no longer take up the main import screen because they normally do not change often.

## Firebird data

When the Windows ODBC DSN connects, V3 uses the same tables as the Excel Power Queries:

- `centralserver` — Stores
- `AGENTTABLE` — RTO agents
- `zonetable` — RTO zones
- `taxzones` — active tax zones and rates
- `dealertable` — Dealer name / Dealer ID / Store
- `Inventory` — existing serial/model/status
- `Customers` — existing names/accounts
- `contracts` — existing RTO contracts

The app does **not** write directly to Firebird. It builds the RTO Pro import files.

## Run on Windows

1. Extract the ZIP into a normal Windows folder.
2. Double-click **`run_windows.bat`**.
3. The app opens at `http://127.0.0.1:5050`.
4. Upload your ShedSuite RTO contracts CSV.
5. Review/edit each contract using the tabs.
6. Click **Save all edits**.
7. Download `RTOProImport.csv` or the Full ZIP.

The first run creates `.venv` and installs the packages automatically.

## Output

- `RTOProImport.csv`
- `Review_Report.csv`
- `Combined_Files\`
- `XML_Files\`
- `RTO_Import_Package.zip`


## V6.3.1 used buildings

The Mapping tab now includes a **Used building** checkbox. When enabled, the building keeps its existing `MODEL1`, while `CONTRACT` receives a unique suffix (`U` first, then `A`, `B`, `C`, etc.) so RTO Pro does not reuse a prior contract number. The app auto-detects likely used buildings from prior RTO Pro contract history and explicit ShedSuite conditions such as Used / Pre-owned / Repo, and displays a visible Used Building tag.


V6.3.1 - Editable Used-Building Suffix
---------------------------------------
- Used buildings still auto-select the first available suffix (U, then A/B/C...).
- The suffix can now be changed manually from a dropdown on the Contract # field.
- Suffixes already used by RTO Pro or another row in the same import are disabled.
- A manually selected suffix is preserved when possible; the pre-import safety check only changes it if it conflicts.


## V6.3.2 manually typed used-building suffix

The Used Building suffix control is now a free-text field. The app still auto-suggests `U`, then `A`, `B`, `C`, etc., but you can type your own alphanumeric suffix manually. The model number remains unchanged, and the resulting contract must be unique and no more than 10 characters total.



## V7 discard + phone normalization

- Each review card has a red **Discard** button. After confirmation, that contract is immediately removed from the current job, its combined PDF is removed, and the RTO CSV/XML/ZIP are rebuilt. Nothing is deleted from ShedSuite itself.
- US phone numbers exported as 11 digits with a leading country code `1` are normalized to the existing `###-###-####` format. The same Primary / Secondary / Ref 1 / Ref 2 comparison rule then runs on the normalized 10 digits.

## V6.3.3 brand/vendor normalization

ShedSuite company-name variations are normalized for RTO Pro BRAND1 and VENDOR1.
The canonical values are ALPINE, 4 SEASONS, GENESIS, SMART SHED, WESTWOOD SHEDS,
LONESTAR SHEDS, TRUE BUILT, YODER STORAGE, and PHOENIX. Legal suffixes and common
variations such as LLC, Buildings LLC, spacing, punctuation, Lone Star, Smart Sheds,
and Four Seasons are accepted automatically. Unknown brands keep their original
upper-cased value.


## V6.3.4–V6.3.5 themes

- Added a persistent normal Dark Mode with a visible Light/Dark toggle.
- Added a hidden Cyberpunk easter-egg theme. It is intentionally absent from the normal theme control.
- Secret activation: press `Alt + Shift + K` anywhere, or click the V7 version label five times quickly on the landing page.
- The selected theme is stored only in the browser (`localStorage`) and follows you across app pages.
- Clicking the normal theme button while Cyberpunk is active returns to normal Dark Mode.
