/**
 * Errors of OOXML package encryption. Each carries a stable `code`: every subpath bundle
 * (`ooxml-core/crypto`, `/xlsx/load`, `/docx/load`, `/pptx`) inlines its own copy of these
 * classes, so a UI should test `error.code` (or `name`) rather than `instanceof` across entries.
 *
 * Moved from `src/pptx/core/utils/ooxml-crypto-errors.ts` (see PROVENANCE.md); `code` and
 * {@link PasswordRequiredError} are new.
 */

/** The stable codes of the package encryption errors. */
export type OoxmlCryptoErrorCode = 'password-required' | 'incorrect-password' | 'data-integrity';

/**
 * The file is an encrypted OOXML package (a compound file holding `EncryptionInfo` and
 * `EncryptedPackage`) and no password was supplied.
 */
export class PasswordRequiredError extends Error {
	public readonly code = 'password-required' as const;

	public constructor(message = 'This file is password protected. Enter the password to open it.') {
		super(message);
		this.name = 'PasswordRequiredError';
	}
}

/**
 * The supplied password did not match the password verifier stored in the encrypted file's
 * EncryptionInfo stream.
 */
export class IncorrectPasswordError extends Error {
	public readonly code = 'incorrect-password' as const;

	public constructor(message = 'The password is incorrect.') {
		super(message);
		this.name = 'IncorrectPasswordError';
	}
}

/**
 * The HMAC computed over the encrypted package did not match the one stored in the
 * EncryptionInfo stream: the file may be corrupted or tampered with.
 */
export class DataIntegrityError extends Error {
	public readonly code = 'data-integrity' as const;

	public constructor(
		message = 'Data integrity check failed. The encrypted file may be corrupted or tampered with.',
	) {
		super(message);
		this.name = 'DataIntegrityError';
	}
}

/** Whether `error` is one of the package encryption errors (by `code`, across bundles). */
export function isOoxmlCryptoError(
	error: unknown,
): error is Error & { readonly code: OoxmlCryptoErrorCode } {
	if (!(error instanceof Error)) return false;
	const code = (error as { code?: unknown }).code;
	return code === 'password-required' || code === 'incorrect-password' || code === 'data-integrity';
}
