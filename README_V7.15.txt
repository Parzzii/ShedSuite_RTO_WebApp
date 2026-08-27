## V7.15 Automatic Delivery Certificate Prefetch

V7.15 starts the ShedSuite Delivery Certificate worker as soon as CSV contracts have been transformed, while the normal direct-link contract/invoice packet is still being assembled. Certificate bytes are cached locally until the base packet is ready, then appended safely without racing the packet writer.

Speed improvements:
- Reuses one authenticated Chromium context/page per ShedSuite login for the whole batch.
- Different ShedSuite company/login accounts prefetch concurrently (default 3, configurable with CERT_PREFETCH_CONCURRENCY=1..5).
- Rows sharing the same login serialize on that one page so the ShedSuite session is not corrupted.
- Existing auth_state cookies are reused across runs, avoiding unnecessary logins when still valid.
- Review status now distinguishes QUEUED, PREFETCHING, PREFETCHED, READY and MISSING.
- PDF-origin contracts remain excluded from ShedSuite certificate lookup.
- Discard/skip still wins over any late prefetch result.

All V7.14 automatic Model/Contract assignment behavior is retained.
