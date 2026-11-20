# Migrating to OOXML 2.0

## Digital signatures

The verifier lives in the `signature` namespace of `ooxml-core/opc` (`import { signature } from 'ooxml-core/opc'`).

A document is shown as signed only when its signature verifies. Every loader reports the verified
state as a `SignatureSummary`, and the old structural helpers are gone.

`ooxml-core/pptx`: `parseSignaturesFromBuffer` is removed and `inspectSignaturesFromBuffer` returns
a summary.

```ts
// 1.x
const signatures = await parseSignaturesFromBuffer(bytes);
const signed = signatures.length > 0;
const intact = (await inspectSignaturesFromBuffer(bytes)).every(
	(s) => s.integrity === 'digestsMatch',
);

// 2.0
import { signature } from 'ooxml-core/opc';
const summary = await inspectSignaturesFromBuffer(bytes); // { state, signatures }
const signed = signature.isVerifiedSignatureState(summary.state); // verified or verified-untrusted
```

Any Office format, from bytes, a JSZip archive or a map of parts:

```ts
import { signature } from 'ooxml-core/opc';
const { state, signatures } = await signature.verifyPackageSignatureSummary(bytes);
```

`ooxml-core/xlsx`: `Workbook.signatures` is a summary.

```ts
// 1.x
if (workbook.signatures) show(`${workbook.signatures.count} signatures`, workbook.signatures.parts);

// 2.0
if (workbook.signatures)
	show(
		workbook.signatures.state,
		workbook.signatures.signatures.map((s) => s.path),
	);
```

`ooxml-core/docx`, `ooxml-core/xlsx`: the English signature warning is no longer in `warnings`.

```ts
// 1.x
const notice = model.warnings.find((warning) => warning.includes('digital signatures'));

// 2.0
import { signatureStateText } from 'ooxml-core/i18n';
const notice = model.signatures && signatureStateText(model.signatures.state, locale)?.description;
```

`ooxml-core/pptx/signature-node`: the canonicalization helpers moved to `ooxml-core/opc`.

```ts
// 1.x
import { canonicalizeNode } from 'ooxml-core/pptx/signature-node';
canonicalizeNode(element, algorithm);

// 2.0
import { signature } from 'ooxml-core/opc';
signature.canonicalize(element, algorithm);
```

## Digital signatures in PowerPoint

Every PowerPoint surface presents a deck's signatures through one descriptor,
`signatureStatusView(summary, locale)` from `ooxml-ui/pptx`. It gives the label, description,
tone, icon and one row per signature in the locale. `view.signed` is true only when the signatures
verify; `view.present` says that signature parts exist (saving removes them).

```ts
// 1.x
const signed = data.hasDigitalSignatures === true;
badge.textContent = signed ? t('pptx.digitalSignatures.badge') : '';

// 2.0
import { signatureStatusView } from 'ooxml-ui/pptx';
const view = signatureStatusView(data.signatures, locale);
badge.hidden = !view.present;
badge.textContent = view.label; // "Signed, signer not verified", "Invalid signature", ...
```

`ooxml-ui/pptx`: the building-block `DigitalSignaturesPort` supplies the summary, and
`SignaturesState` reports the verified state.

```ts
// 1.x
const port = { signed: () => hasSignatures, count: () => signatureCount, content, markDirty };
const { signed, count } = actions.digitalSignatures.state;

// 2.0
const port = { summary: () => data.signatures, content, markDirty };
const { state, signed, present, count } = actions.digitalSignatures.state;
```

`ooxml-ui/pptx`: `HostSignatureInfo` is the core's `SignatureEntry`.

```ts
// 1.x
for (const info of await actions.digitalSignatures.list())
	show(info.signer, info.integrity, info.expired);

// 2.0
for (const entry of await actions.digitalSignatures.list())
	show(entry.signer, entry.state, entry.validTo);
```

`ooxml-ui/pptx`: `loadPresentation` returns the summary.

```ts
// 1.x
const { hasDigitalSignatures, digitalSignatureCount } = await loadPresentation(bytes);

// 2.0
const { signatures } = await loadPresentation(bytes); // SignatureSummary | undefined
const count = signatures?.signatures.length ?? 0;
```

The bindings' signature state moved to the summary in the same way:

```ts
// pptx-react-viewer/internals, 1.x
useLoadContent({ ...input, setHasDigitalSignatures, setDigitalSignatureCount });
// 2.0
useLoadContent({ ...input, setSignatures });

// pptx-vue-viewer/internals, 1.x
const { signaturesApi, hasDigitalSignatures } = useSignatureWorkflow({ signatures, isDirty }); // ParsedSignature[]
// 2.0
const { view } = useSignatureWorkflow({ signatures, locale, isDirty });

// pptx-svelte-viewer, 1.x
loader.hasDigitalSignatures;
loader.digitalSignatureCount;
// 2.0
loader.signatures; // SignatureSummary | undefined

// pptx-angular-viewer, 1.x
inject(SignaturesService);
loadContent.hasDigitalSignatures();
// 2.0
signatureStatusView(loadContent.signatureSummary(), locale);

// pptx-vanilla-viewer, 1.x
store.get().hasDigitalSignatures;
// 2.0
store.get().signatures;
```

`pptx-angular-viewer` also removes the signature helper exports (`isSigned`, `overallStatus`,
`statusLabel` and the rest); read the same answers from `signatureStatusView`.
`pptx-vue-viewer/internals` removes `useSignatures`. `ooxml-ui/pptx` and `pptx-viewer-locales`
drop the unused `pptx.digitalSignatures.*` keys listed in the release notes; signature wording
comes from `signatureStringsFor` in `ooxml-core/i18n`.
