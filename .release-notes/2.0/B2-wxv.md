## Added

- `ooxml-ui`: `<office-ui-signature-status>` shows a document's digital signature state from the `SignatureSummary` a load returns, as a status bar badge (`variant="badge"`), a load notice (`variant="notice"`) or an info panel listing each signer (`variant="panel"`). Its wording comes from the shared `signatureStateText` table in `ooxml-core/i18n`, set by its `locale` property. It emits `office-signature-activate` (badge, View signatures) and `office-signature-dismiss` (notice).
- `ooxml-core`: the shared signature wording gains `view` (View signatures) and `dismiss` in every locale.

## Changed

- Word, Excel and Visio (`ooxml-ui/docx`, `ooxml-ui/xlsx`, `ooxml-ui/visio` and every viewer package built on them) show a signed document's verified state in the status bar, in a notice under the ribbon and on File > Info: Signed, Signed (signer not verified), Invalid signature or Signature not verified, translated in every editor language. A document is shown as signed only when its signature verifies. In a browser no signer is trusted, so a valid signature reads "signer not verified". The notice says that saving changes removes the signatures. Nothing is shown for an unsigned document.

## Security

- A signature that does not verify, or that the verifier cannot check, is no longer presented as a signed document in Word, Excel or Visio.
