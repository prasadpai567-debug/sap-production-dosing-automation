# Redaction checklist

Complete this checklist before publishing any workbook, VBA export, screenshot, or sample document.

- [ ] Remove company and personal author metadata.
- [ ] Remove SharePoint, OneDrive, UNC, and local file paths.
- [ ] Remove real order, material, batch, quantity, Brix, and timestamp values.
- [ ] Remove exception-log entries and live SAP status messages.
- [ ] Remove or abstract SAP plant, order-type, storage-location, transaction, and control identifiers.
- [ ] Remove credentials, tokens, certificates, passwords, and project-protection metadata.
- [ ] Confirm that screenshots contain only synthetic or fully masked data.
- [ ] Confirm that sample PDFs do not contain internal identifiers.
- [ ] Inspect the ZIP contents of any xlsm before release.
- [ ] Review Git history for secrets that may have been committed and later deleted.

The safe public default is documentation plus synthetic examples. Keep the original macro workbook in a private, approved location.
