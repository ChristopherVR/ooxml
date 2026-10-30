// Word 97-2003 binary .doc: the ole2 reader (`readDocModel`) yields the document as a block tree,
// which is converted to a WordprocessingML package (`legacy-doc/`) and opened by the normal .docx
// reader. The loaded document is therefore the same model a .docx gives, and saving it writes a
// .docx through the normal writer; the binary .doc format itself is read-only.
import { readDocFib } from '@christophervr/ole2/ole-document-doc-fib';
import { readDocModel } from '@christophervr/ole2/doc-model';
import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';
import { readCompoundFileStream } from '@christophervr/ole2/ole2-stream-edit';
import type { LoadedDocument } from '../model';
import { loadDocx } from '../parse';
import { buildDocxPackage } from './legacy-doc/package';

export class LegacyDocError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'LegacyDocError';
	}
}

/** The standing notice every loaded .doc carries. */
export const LEGACY_DOC_NOTICE =
	'This Word 97-2003 document was converted when it was opened. Saving writes a Word .docx file; writing back to the binary .doc format is not supported. Revision marks, form-field data, macros, document protection and line numbering are not converted.';

function toBytes(input: Uint8Array | ArrayBuffer): Uint8Array {
	return input instanceof Uint8Array ? input.slice() : new Uint8Array(input.slice(0));
}

/** Load Word 97-2003 binary .doc as a document that saves as .docx. */
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
	return loaded;
}
