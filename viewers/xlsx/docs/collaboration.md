# Collaboration

Real-time co-authoring is **not implemented** for spreadsheets yet. The Word editor (docx-viewer) has a transport-neutral protocol over `ooxml-core/collab`; a spreadsheet equivalent would need operation-level edits in the `xlsx` edit session and conflict rules for structural changes (inserting rows while another person edits a formula that refers to them), and neither exists today.

What works now:

- One person edits a workbook at a time. Listen for `workbook-change` and `dirty-change`, and save with `save()` or by handling the cancelable `file-command` event.
- `readOnly` shows a workbook to many viewers without allowing edits.

This page will describe the protocol once it exists. Until then, nothing on this site claims shared editing.
