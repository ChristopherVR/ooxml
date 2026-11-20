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
