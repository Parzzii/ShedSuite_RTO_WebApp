ShedSuite RTO Web App - V7.20.8
====================================

V7.20.8 - Combined PDF filenames drop the order number
------------------------------------------------------------
- Combined PDF filenames are now just "LAST FIRST" (e.g. "SMITH JOHN.pdf")
  instead of "LAST FIRST orderid" (e.g. "SMITH JOHN 10482.pdf").
- Safety net: this filename is also the actual on-disk key the Delivery
  Certificate, discard, and ZIP export pipeline uses to find each contract's
  PDF (Combined_Files/<name>.pdf). If two customers in the same batch share
  the exact same Last/First name, the second one is automatically named
  "LAST FIRST (2)" so its PDF can't silently overwrite the first customer's.
  This only happens on an actual name collision -- the normal case is
  unaffected and stays exactly "LAST FIRST".
- No other combined-PDF behavior changed.

ShedSuite RTO Web App - V7.20.7
====================================

V7.20.7 - Fixed Magnolia GA/SC/NC/TN picking up stray/out-of-range numbers
------------------------------------------------------------------------------
- Regression from V7.20.4: removing the upper-bound check entirely (to fix
  DBM Phoenix getting stuck at its old ceiling) was too broad. Magnolia
  GA/SC/NC/TN all share the same "0801-" prefix and are only told apart by
  their number range, so their upper bound is a real partition boundary, not
  a stale reference ceiling. With no upper bound at all, an unrelated stray
  value like 0801-2040 in inventory got swept into GA's "used numbers" and
  suggested as the next model, instead of the real last-used 0801-1244
  correctly producing 0801-1245.
- Fix: the upper bound is now enforced ONLY for a prefix shared by more than
  one series (currently just Magnolia's four). A prefix unique to a single
  series (DBM Phoenix, Alpine, Genesis, 4 Seasons, Top Notch, etc.) still has
  no upper bound, so the V7.20.4 fix for real growth past an old ceiling is
  unaffected for all of those.
- Applied consistently across all three places that compute a "last used" or
  "next" model number: build_next_model_suggestions() and normalized_brand()
  callers (rto_transform.py), inventory_profile_payload() (app.py), and
  refresh_model_tracker_from_contracts() (app.py).
- New helpers MODEL_SERIES_SHARED_PREFIXES (rto_transform.py) and
  MODEL_DB_RULES_SHARED_PREFIXES (app.py) compute which prefixes are shared,
  so this stays correct automatically if a new series is ever added that
  happens to reuse an existing prefix.

ShedSuite RTO Web App - V7.20.6
====================================

V7.20.6 - Top Notch is its own Brand/Vendor
-----------------------------------------------
- Correction to V7.20.5: GA orders for "4 Seasons" now set BRAND1/VENDOR1 to
  "TOP NOTCH" instead of "4 SEASONS". WV orders (and any other/blank state)
  are unaffected and still use "4 SEASONS".
- normalized_brand() now takes the order's state and checks it before the
  general Brand/Vendor alias table, so this only affects 4 Seasons/Top Notch
  and every other manufacturer's brand matching is untouched.
- Applied the same brand correction to the Inventory workflow's profile entry
  (INVENTORY_PROFILES['CRF_4_SEASONS_GA']) so Contracts and Inventory modes
  agree on Brand/Vendor for Top Notch.
- Zone (Top Notch, 0155), model series (0155-, starting at 0155-001), Store
  (1 / Carefree), and ShedSuite login remain exactly as set up in V7.20.5.

ShedSuite RTO Web App - V7.20.5
====================================

V7.20.5 - New GA "Top Notch" program for 4 Seasons
-----------------------------------------------------
- ShedSuite uses the same Company Name ("4 Seasons Buildings of WV LLC") for
  both the original WV operation and a new GA operation, but the GA orders
  actually belong to a separate RTO zone/dealer program called "Top Notch".
- New state-aware routing: a Physical Destination State (or Company State /
  Customer Mailing State) of GA now selects:
    - Zone: "Top Notch, 0155" (was "4 Seasons, 0150" for WV)
    - Model series: 0155- (brand new sequence, starts at 0155-001)
  Store (1 / Carefree) and the ShedSuite login stay the same as WV. Any state
  other than GA (including blank) keeps the original WV behavior unchanged.
- Added the same routing to the Inventory workflow's company auto-detect
  (model_series_key) and its Model Series tracker entries, so Inventory mode
  offers "TOP NOTCH (4 Seasons - GA)" as its own trackable series.
- Brand/Vendor for GA/Top Notch contracts is left as "4 SEASONS" (Top Notch is
  treated as the zone/dealer program name, not a different manufacturer).
  This can be changed to "TOP NOTCH" later if desired.
- This state-override pattern (COMPANY_RULES 'state_overrides') is reusable:
  future companies that split into multiple zones by state can be added the
  same way without new code paths.
- Also fixed: refresh_model_tracker_from_contracts() (app.py) had the same
  hardcoded-ceiling bug fixed for model suggestions in V7.20.4, in a third,
  separate code path (the live tracker refresh from RTO Pro contract
  history). It was silently excluding any real contract number above each
  series' old 'max' reference value, so the Model Series tracker panel could
  stop reflecting reality once usage passed that ceiling. Only the floor
  ('min') is enforced now, consistent with the V7.20.4 fix.

ShedSuite RTO Web App - V7.20.4
====================================

V7.20.4 - Model-number suggestion ceiling bug
-----------------------------------------------
- Fixed: next-model suggestions for a series (e.g. DBM Phoenix "0203-") would
  fall back to already-issued numbers (599, 600, 601...) instead of the true
  next number (608) once real inventory grew past that series' original
  upper-bound value (598) copied from Contract_Import.xlsm.
- Root cause: both build_next_model_suggestions() (rto_transform.py) and
  inventory_profile_payload() (app.py) filtered candidate "used" model
  numbers with `lo <= number <= hi`, silently discarding any real inventory
  number above the hardcoded `hi` ceiling for that series.
- Fix: only the lower bound (`lo`) is now enforced when scanning existing
  inventory for a series' highest used number; the upper bound is no longer
  used to exclude legitimate higher numbers. This affects every entry in
  MODEL_SERIES (Alpine, Genesis, Phoenix, 4 Seasons, DBM Phoenix, WWP,
  RAS/YSS, Lonestar, Magnolia GA/SC/NC/TN), not just DBM Phoenix, since all
  of them will eventually grow past their original workbook ceiling too.
- No change to how a series is selected (model_series_key), suffix/used-
  building logic, or fallback behavior when a series has zero inventory.

ShedSuite RTO Web App - V7.20.3
====================================

V7.20.3 - Email review rule refinement
---------------------------------------
- Built directly from the stable V7.20.2 codebase.
- A blank/missing customer email is now allowed and does NOT place the contract in Needs Review.
- A non-empty but malformed/incomplete email is still flagged as Email invalid and places the contract in Needs Review.
- Examples flagged: johnsmith, john@gmail, @gmail.com, john@, or values with spaces.
- Valid addresses such as john@gmail.com remain accepted.
- No other V7.20.2 behavior was changed.

ShedSuite RTO Web App - V7.20.2
====================================

V7.20.2
-------
- Fixed Crestwood Delivery Certificate jobs being skipped when the app found an older Contract_Import.xlsm before the workbook that contained info@whiteriverrto.com.
- Logininfo is now merged from every nearby Contract_Import.xlsm copy; current-app / explicit workbook entries retain priority.
- Crestwood Storage Barns company matching now tolerates LLC/Inc/punctuation/legal-name variations.
- Delivery Certificate errors explicitly say when the mapped credential is absent from all discovered Logininfo tables.

ShedSuite RTO Web App V7.20.1
===========================

This is the single consolidated README / changelog for the project.
New entries are added at the top. Older behavior remains unless a later entry says it was changed.



V7.20.1 - Crestwood / White River direct ShedSuite mapping
-----------------------------------------------------------
- Rolled forward from the stable V7.20 codebase only.
- Added Crestwood Storage Barns as RTO Store/Company 13.
- Forces normal RTO Zone 1301 for Crestwood Storage Barns.
- Maps Crestwood Storage Barns directly to ShedSuite login info@whiteriverrto.com.
- Added common Crestwood Storage Barns LLC name variants.
- Delivery Certificate lookup goes straight to that mapped login when the credential exists in the local Logininfo table.
- No V7.21/V7.22 automatic tenant discovery/account-vault changes are included.
- Password remains in the local Contract_Import.xlsm Logininfo table; it is not stored in this program ZIP.

V7.20 - PDF EPO schedules
--------------------------
- PDF contracts now populate Early Purchase Percentage from provider + RTO term and feed it through the same PAYOFFDISCOUNT conversion already used by ShedSuite CSV imports.
- DBM: 24mo=65%, 36mo=60%, 48mo=55%, 60mo=45%.
- Choice Capital: 24mo=65%, 36mo=60%, 48mo=55%, 60mo=45%.
- RentaBarn / Wolfvalley / Wolf Valley / WVB: 24mo=70%, 36mo=60%, 48mo=50%, 60mo=45%.
- X-Gen: 24mo=70%, 36mo=60%, 48mo=55%, 54mo=55%, 60mo=50%, 72mo=45%.
- Unsupported provider/term combinations are left blank and flagged instead of guessed.
- If a PDF contract term is manually corrected and saved, PAYOFFDISCOUNT is recalculated from the provider schedule.
- The PDF payment audit now shows the detected EPO percentage and resulting PAYOFFDISCOUNT for verification.


V7.19 - Repeat-import Used detection + correction-learning fix
----------------------------------------------------------------
- Repeat imports now identify the physical building by exact Serial before allocating a new model number.
- If that serial is already tied to an RTO Inventory model with rental contract history, the historical model is restored and the row is automatically marked Used.
- Example: first import assigned model/contract 501; running the same CSV again now keeps MODEL1=501 and generates CONTRACT=501U (then 501A, 501B, etc. if needed) instead of assigning 502.
- A STOCK-only serial does not become Used just because it exists in inventory.
- Used-detection reason now explains the exact serial, historical RTO model/status and matching contract history when available.
- Learned Category is now displayed only when learned memory actually changes the category value.
- Learned/manual category text is preserved on save/export instead of being silently remapped back to the automatic category guess.
- Category remains type-or-select; approved values still get canonical spelling, but custom typed values are respected.
- Inventory Category uses the same preserve-typed-value behavior.

V7.18 - Consolidated README + RTO Pro no-open import
-----------------------------------------------------
- Replaces the many per-version README files with this one consolidated README.
- Changelog is maintained newest-to-oldest in a single file named for the current version.
- Customer/contract RTO import now launches with: RTO-win.exe -importcust -noopen app.csv
- Inventory import now launches with: RTO-win.exe -r -noopen inventory_import.csv
- The -noopen switch tells RTO Pro not to display/open the result files after the import completes.
- The generated CSV files are still saved and remain available from the web app when you intentionally want them.

V7.17 - Email quality flag
--------------------------
- Blank customer email is flagged as Email missing.
- Malformed/incomplete values such as johnsmith or john@gmail are flagged as Email invalid.
- Invalid/missing email makes the row appear in Needs Review and highlights the Email field until corrected.
- The app does not guess or auto-complete email domains.

V7.16 - Learn from your corrections
------------------------------------
- Manual corrections saved by the user can become reusable local rules for future imports.
- Learns Store, Dealer, Zone, Category, Brand and Vendor mappings.
- Dealer/Zone learning is Store-aware to reduce cross-store mistakes.
- Category learning ignores dimensions so the same shed style can generalize across sizes.
- Only user-saved corrections teach the system; automatic guesses do not teach themselves.
- Learned rules persist outside the version folder, normally at %LOCALAPPDATA%\ShedSuiteRTO\learned_corrections.json.
- Review cards show a Learned badge when saved local memory changed a field.

V7.15 - Automatic Delivery Certificate prefetch
------------------------------------------------
- Starts ShedSuite Delivery Certificate retrieval as soon as CSV contracts are transformed.
- Certificate retrieval runs while normal direct-link contract/invoice packets are still being assembled.
- Reuses one authenticated Chromium session per ShedSuite login/account.
- Different ShedSuite accounts can prefetch concurrently; default concurrency is 3 and is configurable with CERT_PREFETCH_CONCURRENCY=1..5.
- Rows sharing one login serialize safely on that login session.
- Existing auth state is reused where possible.
- Statuses include QUEUED, PREFETCHING, PREFETCHED, READY and MISSING.
- PDF-origin contracts remain excluded from ShedSuite Delivery Certificate lookup.
- Discard/Skip still wins over a late certificate result.

V7.14 - Automatic batch Model + Contract assignment
----------------------------------------------------
- For normal ShedSuite CSV contracts, reserved Next Model values are placed directly into MODEL1.
- CONTRACT is automatically populated from the same assigned model number.
- Both remain editable.
- Used buildings keep the original physical model and use the suffix workflow instead of consuming a new number.
- PDF contracts keep the PDF Model # and Agreement # / Contract # as authoritative.
- Manual overrides are preserved instead of being overwritten by later recalculation.

V7.13 - Sequential batch model reservation
-------------------------------------------
- Next Model values are reserved across the current batch rather than calculated independently per row.
- Example: if the next model is 501, matching new rows reserve 501, 502, 503, etc.
- Separate companies/model series keep independent counters.
- Used buildings do not consume a new-number slot.
- Sequence recalculates after reorder, discard and Used changes.

V7.12.2 - Large editable Category picker
-----------------------------------------
- Category remains a free-typing field.
- Clicking/focusing it opens a larger readable approved-category suggestion panel.
- Typing filters the list live without blocking manual values.
- Available in Contract Review, All RTO Fields and Inventory.

V7.12.1 - Type-or-select Category
---------------------------------
- Replaced Category dropdown-only behavior with an editable combo field.
- Approved categories still appear as suggestions while the user can type manually.

V7.12 - Approved Category intelligence + color cleanup
------------------------------------------------------
- Uses the approved values from column D of category.xlsx as the known RTO category list.
- Applies spreadsheet-confirmed aliases, keyword rules and conservative fuzzy matching to shed style/description text.
- Inventory description can suggest a category automatically.
- Burnished is normalized to B in generated color text.
- Urethane is removed from generated color text.

V7.11 - PDF reader reliability rebuild
---------------------------------------
- Reworked PDF intake so the parser does not depend on one perfect table encoding.
- Uses multiple extraction strategies including PyMuPDF spatial text, table extraction and alternate PDF text readers.
- Reads important values by their physical label/value positions where possible.
- Avoids manufacturing garbage records from the filename when the PDF text cannot be parsed.
- Recognizes Choice Capital filename patterns such as _CF_ as a fallback provider clue.

V7.10 - PDF Store / Dealer / Zone / Rental Rate correction
-----------------------------------------------------------
- RentaBarn PDF contracts map to Store 3.
- Choice Capital PDF contracts map to Store 12.
- Brand/Vendor still come from the building manufacturer.
- Dealer matching uses normalized/substring/token/fuzzy matching so a source such as Endville Storage can match an RTO dealer such as H&S / Endville Storage.
- Zone selection follows the matched Store/Dealer where possible.
- ZipTax Tax Zone detection remains available.
- PDF Rental Rate uses the printed PMT Before Tax value directly.
- Restored the low-touch OTHER IS... import comment behavior after the V7.9 experiment.

V7.9 - Other-phone SMS opt-out / comment experiment
----------------------------------------------------
- RTO PHONE5 / Other phone exports with CELLOPT2=3 so the Other phone is treated as a cell number opted out of SMS.
- This phone opt-out behavior remains in later versions.
- V7.9 temporarily moved OTHER IS... to a later customer-comment sync; V7.10 restored the prior direct comment workflow for minimum manual steps.

V7.8 - PDF contract/payment mapping fixes
-----------------------------------------
- 90 Days SAC -> SACDATE.
- Agreement # -> CONTRACT.
- Size is normalized compactly, e.g. 10x12.
- Paperless Billing YES -> EMAILINV=1.
- Security Deposit -> EXTRARENT.
- Purchase Reserve -> PAIDDOWN.
- LDW -> GRP.
- PDF PMT Before Tax -> RATE1 directly; no second LDW subtraction.
- Agreement PMT remains 0.00 when item rate is supplied, allowing RTO Pro to calculate the taxed total.
- Total Monthly PMT is retained visibly as a verification reference.
- Safer extraction for multi-column PDFs.

V7.7 - Collapsible + sortable Contract Review
---------------------------------------------
- Click a customer/card header to collapse or expand it.
- Collapse All and Expand All controls.
- Drag handle supports up/down reordering.
- Up/Down buttons provide a non-drag alternative.
- Sort order is persisted and rebuilds RTO CSV/XML/ZIP in that same order.
- Stable hidden row IDs prevent save/discard/background-certificate data from attaching to the wrong customer after sorting.
- Sorting is disabled while filtered/search views are active.

V7.6.2 - Delivery Certificate false-Missing fix
------------------------------------------------
- Restored a long wait for ShedSuite's asynchronously rendered Files list.
- Polls for Delivery Certificate instead of checking immediately after DOM load.
- Refreshes the order page once before declaring a certificate missing.
- Expired/login sessions are reported separately from a genuinely missing certificate.
- Keeps Chromium off-screen while disabling background/occlusion throttling so React keeps rendering.
- Uses broader clickable-element detection when the file name is nested inside a button/link.

V7.6.1 - Delivery Certificate reliability hardening
---------------------------------------------------
- Background worker explicitly selects eligible ShedSuite/CSV contracts.
- Uses a normal headed Chromium context off-screen on Windows rather than relying solely on true headless rendering.
- Broadens Delivery Certificate DOM/text detection.
- Searches more common nested locations for Contract_Import.xlsm / Logininfo credentials.
- PDF-origin contracts remain excluded from ShedSuite certificate lookup.

V7.6 - Mixed CSV + PDF Contract intake
--------------------------------------
- Contracts upload accepts one or many ShedSuite CSV files, supported contract PDFs, or both together.
- CSV and PDF contracts enter the same centralized review/export workflow.
- Duplicate Customer Order / Agreement numbers are kept once; CSV wins when the same contract is already loaded.
- First supported PDF family is the RentaBarn-style detailed contract page; digital/selectable-text PDFs are preferred.
- Extracts common Customer, Unit, Contract, Additional Contact, Employer and Landlord fields.
- New/Used from the PDF can trigger the Used Building workflow automatically.
- Uploaded contract PDF becomes that row's source/combined packet.
- PDF-origin contracts do not launch ShedSuite Delivery Certificate Chromium.
- CSV-origin contracts continue using the background certificate workflow.

V7.5 - Contracts / Inventory home choice
----------------------------------------
- Home screen asks whether the user is entering Contracts or Inventory.
- Contracts opens the existing contract workflow.
- Inventory opens a dedicated inventory-only entry workflow.
- Add Inventory supports one or many buildings.
- Company/model-series selection uses live RTO Inventory data when connected.
- Auto-fills Brand, Vendor, Store and Agent defaults where known.
- Model and Serial are required; duplicate model/serial pairs are blocked in the batch and against live inventory when available.
- Supports Stock, Date Received, Category, Description, Cost, Retail, RTO, Invoice, Rental/Retail, Quantity and BOR.
- Inventory downloads InventoryImport.csv and launches the RTO inventory receiver using -r.
- Known model families include Smart Shed state series, Alpine, Genesis, Phoenix Carefree/Dutch Boy, 4 Seasons, Westwood, Yoder, Lonestar and manual/custom series.

V7.2 - Live Delivery Certificate coordinates
---------------------------------------------
- Re-extracts coordinates as soon as a background Delivery Certificate is appended.
- Review page receives coordinates through background polling without a reload.
- Coordinates are appended to Directions while preserving user-entered Directions text.
- Server-owned coordinate metadata prevents a stale Save All from erasing a newly arrived coordinate.

V7.1 - Multiple CSV input
-------------------------
- Accepts multiple ShedSuite CSV reports in one upload.
- Merges them into one review session.
- Deduplicates overlapping Customer Order IDs; first occurrence wins.
- Keeps each row's own company/dealer/store mapping.
- Review shows how many CSVs were merged and how many duplicates were skipped.
- Original input CSVs are retained only in the local work/job folder for troubleshooting.

V7 - Background Delivery Certificates + Discard + phone normalization
--------------------------------------------------------------------
- Delivery Certificate Chromium moved to a background worker so review/edit can start while certificates download.
- Live queued/downloading/ready/missing status with Retry/Skip behavior.
- Discard removes an unwanted contract from the job and rebuilds generated outputs.
- Background completion merges only server-owned PDF metadata into the latest row, preventing it from overwriting user edits.
- 11-digit US phone numbers beginning with country code 1 are normalized to the underlying 10 digits before the existing phone comparison rule.

V6.3.6 - Discard-ready / phone-country-code groundwork
------------------------------------------------------
- Added immediate contract Discard behavior and rebuilt package outputs after removal.
- Normalized 11-digit US phone values beginning with 1 before comparison/formatting.
- Preserved the existing Primary/Secondary/Ref1/Ref2 selection logic.

V6.3.5 - Dark/Cyberpunk readability + expanded neon palette
------------------------------------------------------------
- Fixed Combined PDFs and RTO Pro Workflow panels that used hard-coded white surfaces and became unreadable in Dark/Cyberpunk themes.
- Made injected panels theme-aware.
- Expanded hidden Cyberpunk accents with pink, red, orange, yellow, cyan and blue neon styling.

V6.3.4 - Light/Dark + hidden Cyberpunk theme
---------------------------------------------
- Added persistent visible Light/Dark theme control.
- Added hidden Cyberpunk easter egg, activated with Alt+Shift+K or repeated version-label clicks on the home page.
- Cyberpunk theme persists between pages and can be exited through the normal theme control.

V6.3.3 - Brand/Vendor normalization
------------------------------------
- BRAND1 and VENDOR1 normalize manufacturer/company variations to exact approved names:
  ALPINE, 4 SEASONS, GENESIS, SMART SHED, WESTWOOD SHEDS, LONESTAR SHEDS, TRUE BUILT, YODER STORAGE, PHOENIX.
- Handles LLC, Buildings, punctuation/spacing and common variants such as Lone Star, Smart Sheds and Four Seasons.
- Unknown manufacturers are preserved/uppercased rather than forced to an incorrect known brand.

V6.3.2 - Free-text Used Building suffix
---------------------------------------
- Used Building suffix changed from fixed dropdown to a free-text box.
- Automatic suggestion still begins with U, then A/B/C... as needed.
- Manual alphanumeric suffixes are accepted and normalized to uppercase.
- MODEL1 remains unchanged; CONTRACT is MODEL1 + suffix.
- Enforces RTO Pro's total Contract length limit and duplicate availability rules.

V6.3.1 - Used Building + editable suffix
----------------------------------------
- Added Used Building checkbox and visible Used badge/detection state.
- Used buildings keep their existing model number.
- Contract receives a unique suffix, beginning with U then A/B/C... as needed.
- Detects likely used/pre-owned/repo buildings from condition/history where available.
- Added manual suffix selection while protecting already-used suffixes.
- Final pre-import check advances a suffix only if the selected contract became unavailable.

V6.3 - Used Building foundation
-------------------------------
- Established the Used Building workflow on top of the V3/V6 mapping system.
- Normal/new-building mapping behavior remains unchanged when Used is not selected.

V6 - V3 base + V5.3 RTO/Delivery Certificate capabilities
----------------------------------------------------------
- Kept the V3 transform/mapping behavior as the authoritative base.
- Added RTO Pro customer/contract upload using a staged Windows CP1252 app.csv and RTO-win.exe -importcust.
- Added Delivery Certificate retrieval using Chromium only for the Delivery Certificate while normal PDFs continue using fast direct ShedSuite URLs.
- Delivery Certificate is appended last to the combined packet.
- Uses the row's V3 ShedSuite login/company mapping to avoid searching unrelated companies.
- Reads legacy Logininfo credentials from Contract_Import.xlsm when available.
- run_windows.bat creates the virtual environment, installs dependencies and Playwright Chromium on first run.

V3 base behavior retained throughout V6/V7
-------------------------------------------
- Review cards organized into Mapping, Inventory, Customer, Tax & Address and PDF/Agent tabs.
- Category/Description kept together for quick editing.
- Customer page shows Mailing and Delivery addresses.
- Existing phone-selection rule:
  1. Primary is always Primary.
  2. Secondary empty -> Ref1.
  3. Secondary matches Primary -> Ref1.
  4. Secondary matches Ref2 -> Ref2.
  5. Secondary matches Ref1 -> Ref1.
  6. Otherwise use Secondary.
- OTHER IS... comment is generated from the chosen reference/secondary source and remains editable.
- ZipTax runs from Delivery Address and suggests an RTO tax zone while allowing manual override.
- Invoice/Salesman/Agent and coordinates are extracted from available PDFs.
- Uses Firebird reference tables for Stores, Agents, Zones, Tax Zones, Dealers, Inventory, Customers and Contracts when connected.
- Outputs include RTOProImport.csv, Review_Report.csv, Combined_Files, XML_Files and the RTO package ZIP.

Running on Windows
------------------
1. Extract the ZIP to a normal writable folder.
2. Double-click run_windows.bat.
3. First run creates .venv and installs Python dependencies / Playwright Chromium.
4. App opens at http://127.0.0.1:5050.
5. Choose Contracts or Inventory and process the batch.

Local files intentionally ignored by Git
----------------------------------------
- .venv/
- venv/
- work/
- __pycache__/
- *.pyc
- .env
