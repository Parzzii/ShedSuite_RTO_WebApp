ShedSuite RTO Web App V7.9

Changes:
1. Other/RTO PHONE5 is always marked as a cell opted out of SMS using CELLOPT2=3 whenever an Other phone exists.
2. Generated OTHER IS <relationship> <name> is removed from the pending application COMMENTS import field.
3. That note is displayed under Directions in the review UI and is synced to the normal RTO Pro CUSTOMERS.COMMENTS field after Refresh RTO Data finds the account.
4. Existing unrelated RTO customer comments are preserved; an older generated OTHER IS segment is replaced rather than duplicated.
