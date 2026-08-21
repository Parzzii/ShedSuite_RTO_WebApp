ShedSuite RTO Web App V6.3
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


V6.3 - Used Buildings
---------------------
- Adds a Used building checkbox to every contract card.
- When checked, MODEL1 stays unchanged and CONTRACT is automatically suffixed.
- Suffix order is U first, then A, B, C ... Z (without repeating U).
- The generated contract must remain within RTO Pro's 10-character CONTRACT limit.
- Auto-detects used buildings when the MODEL1 has prior RTO Pro contract history,
  or when the ShedSuite Condition explicitly indicates used/pre-owned/repo.
- Shows a USED BUILDING / USED DETECTED tag in the review card.
- Keeps the normal V3/V6 mapping behavior untouched for non-used buildings.
