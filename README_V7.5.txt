ShedSuite RTO Web App V7.5
===========================

NEW WORKFLOW CHOICE
- Home screen asks: Contracts or Inventory.
- Contracts = the existing V7.2 workflow.
- Inventory = dedicated RTO Pro inventory-only receiving.

INVENTORY MODE
- Click Add inventory to add one or many buildings.
- Choose Company / Model Series on each item.
- Uses the live RTO Pro Inventory table for the next model suggestion when Firebird is connected.
- Multiple new items for the same company automatically advance the suggestion within the current batch.
- Auto-fills canonical Brand + Vendor, Store and Agent from existing rules.
- Model and Serial are required.
- V7.5 also requires Company and Store for this Central Server workflow.
- Supports Stock, Date Received, Category, Description, Cost, Retail, RTO price, Invoice, Rental/Retail, Quantity and BOR.
- Duplicate model + serial combinations are blocked when found in the current batch or live RTO inventory.
- Downloads InventoryImport.csv.
- "Import Inventory to RTO Pro" launches the RTO inventory receiver using -r.

MODEL SERIES
- SMART SHED — GA / SC / NC / TN
- ALPINE
- GENESIS
- PHOENIX — Carefree
- 4 SEASONS
- PHOENIX — Dutch Boy
- WESTWOOD SHEDS
- YODER STORAGE
- LONESTAR SHEDS
- TRUE BUILT — manual model (no numbering series exists in the original workbook, so V7.5 does not invent one)
- OTHER / CUSTOM

All V7.2 contract functionality remains available from the Contracts choice.
