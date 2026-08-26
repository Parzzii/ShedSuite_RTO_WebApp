ShedSuite RTO Web App V7.8

PDF contract mapping fixes:
- 90 Days SAC -> exact SACDATE printed in PDF
- Agreement # -> CONTRACT
- compact size e.g. 10x12
- Paperless Billing YES -> EMAILINV 1
- Security Deposit -> EXTRARENT
- Purchase Reserve -> PAIDDOWN
- LDW -> GRP
- PMT Before Tax -> RATE1 directly (no second LDW subtraction)
- PMT = 0.00 for PDF inventory contracts so RTO Pro calculates the taxed total
- Total Monthly PMT retained visibly for verification
- safer multi-column PDF table extraction
