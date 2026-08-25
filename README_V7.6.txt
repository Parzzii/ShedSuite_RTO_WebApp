ShedSuite RTO Web App V7.6
===========================

V7.6 adds mixed contract intake on top of V7.5.

CONTRACT INPUTS
---------------
- Upload one or many ShedSuite CSV files.
- Upload one or many supported RentaBarn-style contract PDFs.
- CSV and PDF files can be mixed in the same batch.
- Duplicate Customer Order / Agreement numbers are kept only once; CSV wins when already loaded.
- Contract PDFs are parsed into the same internal ShedSuite-style source schema, so the existing RTO transform/review/export logic remains centralized.

PDF CONTRACT PARSER
-------------------
The parser reads digital/text-based RentaBarn-style PDFs and extracts the common sections:
- Customer Information
- Unit Information
- Contract Information
- Co-Renter
- Additional Contacts
- Employer / Source of Income
- Landlord

Important mapped values include customer name/address/phones/DOB/SSN/DL/email, manufacturer, serial, model, New/Used, style/type/size/colors, cash price, dealer, salesperson, agreement, contract date, RTO term, taxes, LDW, payment amounts, references and employer data.

The PDF's Manufacturer is normalized into existing company rules when it matches known families such as Alpine, 4 Seasons, Genesis, Smart Shed, Westwood, Lonestar, Yoder or Phoenix. Unknown manufacturers are preserved for manual review.

USED BUILDINGS
--------------
If the PDF says New or Used = USED, the existing V7 used-building workflow automatically turns on and keeps MODEL1 unchanged while applying the normal unique contract suffix rule.

PDF PACKETS / DELIVERY CERTIFICATES
-----------------------------------
- The uploaded contract PDF itself becomes that row's combined source PDF.
- PDF-source contracts do not run ShedSuite Chromium / Delivery Certificate lookup.
- ShedSuite CSV rows still use the existing fast direct packet + background headless Delivery Certificate workflow.
- Mixed CSV/PDF batches can be edited together while ShedSuite certificates continue in the background.

LIMITATION
----------
The first supported PDF format is the RentaBarn-style table layout discussed during development. The PDF must contain selectable text. Scanned-image PDFs/screenshots are intentionally not OCR'd; re-download the original digital PDF for reliable parsing.

Everything from V7.5 remains, including Inventory-only receiving, Firebird lookups, model suggestions, used-building suffix editing, phone normalization, discard, brand/vendor normalization, dark mode, hidden Cyberpunk mode, multi-CSV input, and live background coordinate updates.
