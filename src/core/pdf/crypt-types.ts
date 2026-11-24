/**
 * What an encrypted PDF reports, kept apart from the security handler (`crypt.ts`) so a reader
 * of unencrypted files never loads the ciphers and digests.
 */
export class PdfPasswordError extends Error {
	constructor(
		message: string,
		readonly code: 'password-required' | 'incorrect-password',
	) {
		super(message);
	}
}

/** What the file's permission flags allow (ISO 32000-1 table 22). */
export interface PdfPermissions {
	print: boolean;
	modify: boolean;
	copy: boolean;
	annotate: boolean;
	fillForms: boolean;
	extractForAccessibility: boolean;
	assemble: boolean;
	printHighQuality: boolean;
}
