ShedSuite RTO Web App V7.10

PDF mapping + low-touch workflow correction

Changes from V7.9:
- Restored the V7.8-style OTHER IS... comment in the normal import COMMENTS field and review Comment box.
- Kept PHONE5 / Other phone as CELLOPT2=3 (cell number opted out of SMS).
- PDF provider mapping: RentaBarn = Store 3; Choice Capital = Store 12.
- Brand/Vendor still come from the building manufacturer, not the PDF provider.
- Smarter PDF Dealer matching against the live Firebird dealer list: exact, normalized, substring/token and guarded fuzzy matching, always preferring the selected store.
- Example: PDF "Endville Storage" can match RTO dealer "H&S / Endville Storage".
- PDF normal Zone automatically matches the selected store/dealer where possible; if a store has only one zone, that unique zone is selected.
- Existing ZipTax automatic Tax Zone logic remains active; an exact PDF Tax Code is also accepted when it matches an active RTO tax-zone code.
- PDF Rental Rate / RATE1 now uses the value physically printed to the right of "PMT Before Tax" as the authoritative value. It is never reduced by LDW on the PDF path.
- Security Deposit, Purchase Reserve, LDW, Total Tax and Total Monthly PMT also use position-aware extraction to reduce side-by-side table number drift.
- Refresh RTO Data no longer performs a separate comment sync step.

Notes:
- Dealer and Zone auto-matching require a successful Firebird connection because those lists come from RTO Pro. Store 3/12 defaults work even when Firebird is unavailable.
