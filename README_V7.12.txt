ShedSuite RTO Web App V7.12

CATEGORY RULES
- Uses only the non-empty canonical categories from category.xlsx D1:D170.
- 28 approved values are embedded in category_rules.py.
- The green/confirmed workbook mappings are embedded as exact aliases.
- For unknown descriptions, the app uses high-confidence semantic rules, then a conservative fuzzy match. Weak matches fall back to Utility rather than inventing an unrelated category.
- Contract review and Inventory-only Category controls are dropdowns limited to the approved list.
- Inventory descriptions can auto-suggest the category.

COLOR RULES
- Burnished -> B (example: Burnished Slate -> B Slate).
- Urethane is removed wherever it appears (example: Urethane Burnished Slate -> B Slate).
- The rules apply to generated contract descriptions and saved inventory descriptions.

All V7.11 PDF reader, V7.10 PDF mapping, background Delivery Certificate, sorting/collapse, multi-CSV, used-building, and RTO Pro workflows are retained.
