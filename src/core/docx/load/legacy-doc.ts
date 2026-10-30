// Word 97-2003 binary .doc: the ole2 reader (`readDocModel`) yields the document as a block tree,
// which is converted to a WordprocessingML package (`legacy-doc/`) and opened by the normal .docx
// reader. The loaded document is therefore the same model a .docx gives. `save()` keeps the
// binary .doc format (unchanged bytes, or paragraph text edits); `saveDocx()` writes the full
// model as a .docx through the normal writer.
import { readDocFib } from '@christophervr/ole2/ole-document-doc-fib';
import { readDocModel } from '@christophervr/ole2/doc-model';
import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';
import { readCompoundFileStream } from '@christophervr/ole2/ole2-stream-edit';
import type { LoadedDocument } from '../model';
import { loadDocx } from '../parse';
import { binaryDocSaver } from './legacy-doc/binary-save';
import { LegacyDocError } from './legacy-doc/error';
import { buildDocxPackage } from './legacy-doc/package';

export { LegacyDocError };

/** The standing notice every loaded .doc carries. */
export const LEGACY_DOC_NOTICE =
	'This Word 97-2003 document was converted when it was opened. Save it as a Word .docx file to keep every change; the .doc format itself accepts paragraph text edits only. Revision marks, form-field data, macros, document protection and line numbering are not converted.';

function toBytes(input: Uint8Array | ArrayBuffer): Uint8Array {
	return input instanceof Uint8Array ? input.slice() : new Uint8Array(input.slice(0));
}

/** Load Word 97-2003 binary .doc. `save()` keeps the .doc format; `saveDocx()` saves the full model. */
export async function loadLegacyDoc(input: Uint8Array | ArrayBuffer): Promise<LoadedDocument> {
	const original = toBytes(input);
	let streams: string[] = [];
	let wordDocument: Uint8Array | null = null;
	try {
		const parsed = parseOle2(
			original.buffer.slice(
				original.byteOffset,
				original.byteOffset + original.byteLength,
			) as ArrayBuffer,
		);
		streams = parsed.entries.map((entry) => entry.name);
		wordDocument = readCompoundFileStream(original, ['WordDocument']) ?? null;
	} catch (error) {
		throw new LegacyDocError(
			`Invalid or truncated OLE compound file: ${error instanceof Error ? error.message : 'parse failed'}`,
		);
	}
	if (streams.includes('EncryptionInfo') || streams.includes('EncryptedPackage')) {
		throw new LegacyDocError(
			'This is a password-protected .docx, not a legacy .doc: open it with loadDocument and its password.',
		);
	}
	if (wordDocument) {
		try {
			readDocFib(wordDocument);
		} catch (error) {
			// The shared FIB reader rejects encryption before stream unwrapping.
			if (error instanceof Error && error.message === 'Encrypted or obfuscated Word document')
				throw new LegacyDocError('Encrypted legacy Word documents are not supported.');
			// Other malformed FIBs retain the existing unsupported-document diagnostic below.
		}
	}
	let docModel;
	try {
		docModel = readDocModel(original);
	} catch (error) {
		throw new LegacyDocError(
			`The legacy Word document is too large or too complex to open: ${error instanceof Error ? error.message : 'limit exceeded'}`,
		);
	}
	if (!docModel) {
		throw new LegacyDocError(
			streams.includes('WordDocument')
				? 'Unsupported or malformed legacy Word document.'
				: 'OLE file does not contain a legacy Word document.',
		);
	}
	const converted = await buildDocxPackage(docModel);
	const loaded = await loadDocx(converted.bytes);
	loaded.model.warnings.unshift(LEGACY_DOC_NOTICE);
	for (const warning of converted.warnings) loaded.model.warnings.push(warning);
	for (const warning of docModel.warnings) loaded.model.warnings.push(`Legacy .doc: ${warning}`);
	const saveBinary = binaryDocSaver(original, loaded.model);
	return {
		...loaded,
		async save(model = loaded.model) {
			return saveBinary(model);
		},
		saveDocx: (model = loaded.model, pendingMedia, options) =>
			loaded.save(model, pendingMedia, options),
	};
}
