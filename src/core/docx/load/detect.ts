import { PasswordRequiredError } from '../../crypto/errors.js';
import { decryptOoxmlPackage, isEncryptedOoxmlPackage } from '../../crypto/index.js';
import type { LoadedDocument } from '../model.js';
import { loadDocx } from '../parse.js';
import { loadLegacyDoc } from './legacy-doc.js';

/**
 * A detected document format. `'encrypted'` is a .docx saved with a password to open (a compound
 * file holding EncryptionInfo and EncryptedPackage), not a Word 97-2003 .doc.
 */
export type DocumentFormat = 'docx' | 'doc' | 'encrypted';

const isZip = (bytes: Uint8Array): boolean =>
	bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;

export function detectDocumentFormat(input: Uint8Array | ArrayBuffer): DocumentFormat {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	if (
		[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((value, index) => bytes[index] === value)
	)
		return isEncryptedOoxmlPackage(bytes) ? 'encrypted' : 'doc';
	if (isZip(bytes)) return 'docx';
	throw new Error('Unsupported file. Open a DOCX or Word 97-2003 DOC document.');
}

export interface LoadDocumentOptions {
	/**
	 * The password to open an encrypted (password-protected) .docx. Without it such a file throws
	 * `PasswordRequiredError` (`code: 'password-required'`); a wrong one throws
	 * `IncorrectPasswordError` (`code: 'incorrect-password'`).
	 */
	password?: string;
}

const ENCRYPTED_WARNING =
	'This document was opened with a password. Saving writes it without the password; encrypt the saved bytes again to keep it protected.';

/** Sniff content, never trust a user-supplied filename extension. */
export async function loadDocument(
	input: Uint8Array | ArrayBuffer,
	options: LoadDocumentOptions = {},
): Promise<LoadedDocument> {
	const format = detectDocumentFormat(input);
	if (format === 'doc') return loadLegacyDoc(input);
	if (format === 'docx') return loadDocx(input);
	if (options.password === undefined) {
		throw new PasswordRequiredError(
			'This document is password protected. Enter the password to open it.',
		);
	}
	const plain = await decryptOoxmlPackage(input, options.password);
	if (!isZip(plain)) throw new Error('The decrypted file is not a Word document package.');
	const loaded = await loadDocx(plain);
	loaded.model.warnings.push(ENCRYPTED_WARNING);
	return loaded;
}
