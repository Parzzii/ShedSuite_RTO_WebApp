ShedSuite RTO Web App V7.11

PDF reader reliability rebuild

- Adds a geometry-based PDF reader that reads values from the same printed row as each label inside Customer / Unit / Contract sections.
- Keeps left and right halves of the detailed-information page separate so values cannot drift across columns.
- Uses multiple text engines: PyMuPDF, PyMuPDF blocks, pypdf and pdfplumber.
- PyMuPDF table extraction and pdfplumber table extraction remain as fallbacks.
- The strict single-layout gate from V7.10 is removed. A PDF can be accepted when recognizable form fields are found even if one engine produces unusual text order.
- Choice Capital filename token _CF_ is recognized as a provider fallback when the logo/provider name is artwork rather than selectable text.
- Agreement #, Model # and Serial # no longer accidentally keep the printed # character as the value.
- V7.10 store/dealer/zone/payment/comment/phone logic is retained.

If a PDF truly contains no extractable text, V7.11 now rejects it with a clear message instead of creating an Unnamed / Not Found contract from the filename.
