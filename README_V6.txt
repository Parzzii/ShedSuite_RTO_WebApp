ShedSuite RTO Web App V7.1
========================

Base: V3.0.0

V6 intentionally keeps V3 mapping logic. rto_transform.py is unchanged from the supplied V3.

Added from V5.3:
1. RTO Pro upload/import workflow
   - stages a Windows ANSI/CP1252 app.csv
   - tries writable locations automatically
   - launches: RTO-win.exe -importcust app.csv
   - writes diagnostic launched CSV / CMD / command text into the job package

2. Delivery Certificate download
   - normal PDFs use V3 direct CSV URLs (fast)
   - Chromium is used ONLY for Delivery Certificate
   - Delivery Certificate is appended last
   - V3 row _login is authoritative, preventing cross-company login searches for known companies
   - credentials are read locally from the old Contract_Import.xlsm Logininfo table

Startup:
- Run run_windows.bat
- On first run it installs Python dependencies and Playwright Chromium.
- .env is created from .env.example if missing.

Important:
- If Contract_Import.xlsm cannot be auto-found, set LEGACY_XLSM_PATH in .env.
- RTO_EXE defaults to C:\RTOwin\RTO-win.exe.


V6.3.1 - Used Buildings
---------------------
- Adds a Used building checkbox to every contract card.
- When checked, MODEL1 stays unchanged and CONTRACT is automatically suffixed.
- Suffix order is U first, then A, B, C ... Z (without repeating U).
- The generated contract must remain within RTO Pro's 10-character CONTRACT limit.
- Auto-detects used buildings when the MODEL1 has prior RTO Pro contract history,
  or when the ShedSuite Condition explicitly indicates used/pre-owned/repo.
- Shows a USED BUILDING / USED DETECTED tag in the review card.
- Keeps the normal V3/V6 mapping behavior untouched for non-used buildings.


V6.3.1 - Editable Used-Building Suffix
---------------------------------------
- Used buildings still auto-select the first available suffix (U, then A/B/C...).
- The suffix can now be changed manually from a dropdown on the Contract # field.
- Suffixes already used by RTO Pro or another row in the same import are disabled.
- A manually selected suffix is preserved when possible; the pre-import safety check only changes it if it conflicts.


V6.3.2 - Manually Typed Used-Building Suffix
------------------------------------------------
- Replaces the fixed suffix dropdown with a free-text suffix box.
- Auto-detection still suggests U first, then A/B/C... when a building is marked used.
- You can type a different suffix manually; letters and numbers are accepted and normalized to uppercase.
- MODEL1 remains unchanged; CONTRACT is always MODEL1 + the typed suffix.
- The final contract must fit RTO Pro's 10-character CONTRACT limit and must not already exist.
- A manual suffix is preserved through RTO Pro import when it is still available.


V6.3.3 - Brand/Vendor Normalization
-----------------------------------
BRAND1 and VENDOR1 now normalize ShedSuite manufacturer/company-name variants to
these exact RTO Pro values: ALPINE, 4 SEASONS, GENESIS, SMART SHED, WESTWOOD SHEDS,
LONESTAR SHEDS, TRUE BUILT, YODER STORAGE, and PHOENIX. LLC/Buildings/punctuation/
spacing and common variants such as Lone Star, Smart Sheds, and Four Seasons are
recognized automatically.


V6.3.4/6.3.5 - Themes
----------------
- Visible persistent Light/Dark mode toggle.
- Hidden Cyberpunk theme easter egg.
- Secret: Alt+Shift+K, or click the version label five times quickly on the first page.



V7.1 - Multiple CSV Input
-------------------------
- Select one or multiple ShedSuite RTO contracts CSVs on the first page.
- All selected CSVs merge into one review session.
- Overlapping Customer Order IDs are skipped so the same contract is not imported twice.
- The review bar shows how many CSVs were merged and how many duplicates were skipped.
- Original CSVs are kept only in the local work/job Input_CSVs folder.

V7 - Discard + Phone Normalization
----------------------------------------
- Review cards now include a Discard button that immediately removes the selected contract from the current job and rebuilds RTO CSV/XML/ZIP output.
- The discarded contract's combined PDF is removed from the package as well.
- 11-digit US phone numbers beginning with country code 1 are reduced to the underlying 10 digits and formatted ###-###-#### before the existing phone comparison rule is applied.


V7 - Background Delivery Certificates
-------------------------------------
Delivery Certificate Chromium now runs headlessly in a background worker.
You can review, edit, discard, and save contracts while certificates download.
The review page shows live progress and supports Retry/Skip for certificate jobs.
