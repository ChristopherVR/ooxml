## Breaking

- `ooxml-ui/pptx`: the building-block `DigitalSignaturesPort` supplies `summary()` (the deck's `SignatureSummary`) instead of `signed()` and `count()`. `describeSignaturesState` takes a `SignatureSummary`, and `SignaturesState` gains `state`, `signed` and `present`.
- `ooxml-ui/pptx`: `HostSignatureInfo` is the core's `SignatureEntry` (`path`, `state`, `signer`, `issuer`, `signingTime`, `validFrom`, `validTo`). The `expired` and `integrity` fields are gone.
- `ooxml-ui/pptx`: `loadPresentation` returns `signatures` (a `SignatureSummary`) instead of `hasDigitalSignatures` and `digitalSignatureCount`.
- `ooxml-ui/pptx`, `pptx-viewer-locales`: the dictionaries no longer contain the unused keys `pptx.digitalSignatures.badge`, `badgeSigned`, `badgeTitle`, `headerSigned`, `invalidHeader`, `signed`, `statusValid`, `statusInvalid`, `statusExpired`, `statusUnknownCA`, `statusUnverified`, `serial` and `certificate`. Signature wording now comes from `signatureStringsFor` in `ooxml-core/i18n`.
- `pptx-react-viewer`: `pptx-react-viewer/internals` changed. `UseLoadContentInput` takes `setSignatures` instead of `setHasDigitalSignatures` and `setDigitalSignatureCount`, and `UseViewerDialogsInput` takes `signatures` instead of `hasDigitalSignatures`.
- `pptx-vue-viewer`: `pptx-vue-viewer/internals` no longer exports `useSignatures`, `UseSignaturesResult` or `OverallSignatureStatus`. `useSignatureWorkflow` takes a `SignatureSummary` ref and a locale and returns `view` instead of `signaturesApi` and `hasDigitalSignatures`. `useLoadContent` returns `signatures` as a `SignatureSummary` and no longer exports `parseSignaturesFromBuffer`.
- `pptx-svelte-viewer`: `PresentationLoader` has `signatures` (a `SignatureSummary`) instead of `hasDigitalSignatures` and `digitalSignatureCount`, and `Ribbon` takes `signatureView` instead of those two props.
- `pptx-angular-viewer`: `SignaturesService` and the signature helper exports (`isSigned`, `worstStatus`, `overallStatus`, `headerLabel`, `statusLabel`, `statusKind`, `signerName`, `signatureTimestamp`, `signatureKey`, `signatureCountLabel`, `OverallSignatureStatus`, `SignatureStatusKind`) are removed. `SignaturesPanelComponent` takes `summary` instead of `signatures`, and `LoadContentService` has `signatureSummary` instead of `signatures`, `hasDigitalSignatures` and `digitalSignatureCount`.
- `pptx-vanilla-viewer`: `ViewerState` has `signatures` (a `SignatureSummary`) instead of `hasDigitalSignatures` and `digitalSignatureCount`.

## Migration

- PowerPoint viewers: read `PptxData.signatures` and present it with `signatureStatusView(summary, locale)` from `ooxml-ui/pptx`. It gives the label, description, tone, icon and one row per signature, worded in the locale. Show a deck as signed only when `view.signed` is true. `view.present` says that signature parts exist, which is what the save warning needs.
- `ooxml-ui/pptx`: a host implementing `DigitalSignaturesPort` returns the loaded deck's `SignatureSummary` from `summary()` in place of `signed()` and `count()`.
- `ooxml-ui/pptx`, `pptx-viewer-locales`: a host that translated one of the removed `pptx.digitalSignatures.*` keys can drop it; use `signatureStringsFor(locale)` from `ooxml-core/i18n` for signature wording.

## Added

- `ooxml-ui/pptx`: `signatureStatusView`, with the `SignatureStatusView`, `SignatureStatusRow`, `SignatureTone` and `SignatureIcon` types, describes a deck's signatures for every badge, panel and notice.
- `ooxml-ui/pptx`: `signatureLoadNotice(view)` and `SIGNATURE_NOTICE_ID` give the notice a binding shows when a signed deck opens. `CompatibilityWarningToast` and the `pptx-ui-compat-toasts` element accept an already translated `text`.
- PowerPoint viewers: a signed deck opens with a dismissible notice of its signature state and the note that saving changes removes the signatures, translated in every UI language. It returns when another deck loads.
- `pptx-vanilla-viewer`: a signature badge under the ribbon that opens the Digital Signatures dialog, as the other bindings have.

## Changed

- PowerPoint viewers: the badge, the Digital Signatures panel and its rows show the verified state of each signature: Signed, Signed (signer not verified), Invalid signature or Signature not verified. A browser checks the signed parts' digests, the signature value and the signer certificate's key but trusts no signer, so a valid signature reads "Signed, signer not verified" at best.

## Fixed

- PowerPoint viewers: a deck with an invalid, forged or unverifiable signature is no longer labelled "Signed". Every binding used to show "Signed" whenever the deck contained a signature part.
