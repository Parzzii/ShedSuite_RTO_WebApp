## V7.16 Learn From Your Corrections

V7.16 adds conservative local correction learning to the Contracts review workflow. When you manually correct a repeatable mapping and click Save all edits, the app remembers the source-to-RTO mapping and can apply it automatically on future imports.

Learns:
- Source company/provider/dealer -> Store
- Source dealer + Store -> Dealer ID / dealer name
- Source dealer + Store + Dealer -> Zone
- Building model/style/description -> Category (sizes are ignored so the same style can generalize)
- Source manufacturer/company -> Brand and Vendor

Safety:
- Only user-saved changes become rules; automatic guesses do not teach themselves.
- Exact/high-confidence matches are preferred. Fuzzy reuse is deliberately conservative.
- Dealer/Zone rules are scoped to Store so similar dealer names in different stores do not cross-map.
- A later correction to the same source updates the learned target, so mistakes can be taught over.
- Learned rules are stored outside the version folder so upgrading ZIPs does not erase them. On Windows the default is %LOCALAPPDATA%\ShedSuiteRTO\learned_corrections.json.
- Rows that used memory show a "Learned: ..." badge on the review card.

V7.15 Delivery Certificate prefetch and all earlier V7.14/V7.13 behavior remain in place.
