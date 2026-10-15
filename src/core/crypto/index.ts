// OOXML package encryption ([MS-OFFCRYPTO]), shared by every format: a password-to-open .xlsx,
// .docx or .pptx is a compound file (CFB, from ole2, inlined at build time) holding an
// `EncryptionInfo` stream and the AES-encrypted zip in `EncryptedPackage`. Agile (Office 2010+)
// and Standard (Office 2007) are decrypted; agile is written by default.
//
// Moved from the pptx area (`src/core/pptx/core/utils/ooxml-crypto*.ts`, `encryption-detection.ts`;
// see PROVENANCE.md). Not re-exported from the root entry: it pulls the CFB codec in.
import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';
export { createOfficeGuid } from './uuid';
import {
	decryptAgilePackage,
	decryptStandardPackage,
	verifyAgileDataIntegrity,
	verifyAgilePassword,
	verifyStandardPassword,
	type DecryptOptions,
} from './decrypt';
import { openEncryptedPackage, toArrayBuffer, toBytes } from './detect';
import { encryptAgile, encryptStandard } from './encrypt';
import { parseEncryptionInfo } from './encryption-info';
import { IncorrectPasswordError } from './errors';
import type { EncryptionInfo, EncryptionOptions, StandardEncryptionInfo } from './types';

export {
	DataIntegrityError,
	IncorrectPasswordError,
	PasswordRequiredError,
	isOoxmlCryptoError,
	type OoxmlCryptoErrorCode,
} from './errors';
export {
	EncryptedFileError,
	detectFileFormat,
	isCompoundFile,
	isEncryptedOoxmlPackage,
	openEncryptedPackage,
	type EncryptedPackageStreams,
	type FileFormatDetection,
} from './detect';
export type {
	EncryptionAlgorithm,
	EncryptionInfo,
	EncryptionOptions,
	EncryptionScheme,
	StandardEncryptionInfo,
} from './types';
export { RC4_ALG_ID, parseEncryptionInfo } from './encryption-info';
export {
	BLOCK_KEYS,
	computeStandardKeyBase,
	deriveAgileKey,
	deriveStandardKey,
	deriveStandardKeyFromBase,
	generateIV,
} from './key-derivation';
export {
	decryptAgilePackage,
	decryptStandardPackage,
	verifyAgileDataIntegrity,
	verifyAgilePassword,
	verifyStandardPassword,
	type DecryptOptions,
} from './decrypt';
export { encryptAgilePackage, encryptStandardPackage } from './encrypt';
export {
	buildAgileEncryptionInfoXml,
	buildEncryptionInfoStream,
	buildStandardEncryptionInfoStream,
} from './encryption-info-write';

const isStandard = (
	info: EncryptionInfo | StandardEncryptionInfo,
): info is StandardEncryptionInfo => 'isStandard' in info && info.isStandard;

/** The two streams, or a descriptive error naming the missing one. */
function streams(bytes: Uint8Array): { info: Uint8Array; data: Uint8Array } {
	const opened = openEncryptedPackage(bytes);
	if (opened) return { info: opened.encryptionInfo, data: opened.encryptedPackage };
	// Re-open to report which stream is missing (or let the compound-file parser explain).
	const file = parseOle2(toArrayBuffer(bytes));
	if (!file.getStream('EncryptionInfo')) {
		throw new Error(
			'EncryptionInfo stream not found. The file may not be an encrypted OOXML package.',
		);
	}
	throw new Error('EncryptedPackage stream not found. The file may be corrupted.');
}

/**
 * Decrypts a password-protected OOXML package (agile or Standard encryption) to the plain zip.
 *
 * @throws IncorrectPasswordError (`code: 'incorrect-password'`) when the password is wrong.
 * @throws DataIntegrityError (`code: 'data-integrity'`) when the agile HMAC does not match or
 *   is absent (pass `allowMissingIntegrity: true` to accept a package without one).
 *   Standard encryption has no integrity record at all: see {@link inspectOoxmlEncryption}.
 * @throws Error when the bytes are not an encrypted package or the scheme is unsupported.
 */
export async function decryptOoxmlPackage(
	input: Uint8Array | ArrayBuffer,
	password: string,
	options: DecryptOptions = {},
): Promise<Uint8Array> {
	const { info: infoStream, data } = streams(toBytes(input));
	const info = parseEncryptionInfo(infoStream);
	if (isStandard(info)) {
		const key = await verifyStandardPassword(info, password);
		if (!key) throw new IncorrectPasswordError();
		return new Uint8Array(await decryptStandardPackage(data, key));
	}
	const key = await verifyAgilePassword(info, password, options);
	if (!key) throw new IncorrectPasswordError();
	await verifyAgileDataIntegrity(info, key, data, options);
	return new Uint8Array(await decryptAgilePackage(data, key, info));
}

/** What an encrypted package offers, from {@link inspectOoxmlEncryption}. */
export interface OoxmlEncryptionSummary {
	scheme: 'agile' | 'standard';
	/**
	 * Whether the ciphertext is authenticated. False for Standard encryption (no HMAC exists in
	 * the scheme, so a decrypted package may have been modified) and for an agile package whose
	 * `dataIntegrity` was stripped.
	 */
	integrityProtected: boolean;
}

/** The scheme of an encrypted package and whether it is integrity protected; no password needed. */
export function inspectOoxmlEncryption(input: Uint8Array | ArrayBuffer): OoxmlEncryptionSummary {
	const info = parseEncryptionInfo(streams(toBytes(input)).info);
	if (isStandard(info)) return { scheme: 'standard', integrityProtected: false };
	return { scheme: 'agile', integrityProtected: info.dataIntegrity !== undefined };
}

/**
 * Encrypts an OOXML package (the zip bytes of a .xlsx, .docx, .pptx, ...) with a password to
 * open. Writes the agile scheme (AES-256, SHA-512, 100,000 spins: what Office 2010+ writes) by
 * default; `encryptionScheme: 'standard'` writes the Office 2007 scheme.
 */
export function encryptOoxmlPackage(
	input: Uint8Array | ArrayBuffer,
	password: string,
	options: EncryptionOptions = {},
): Promise<Uint8Array> {
	const bytes = toBytes(input);
	return options.encryptionScheme === 'standard'
		? encryptStandard(bytes, password, options)
		: encryptAgile(bytes, password, options);
}

/** Whether `password` opens the encrypted package, without decrypting it. Never throws. */
export async function verifyOoxmlPackagePassword(
	input: Uint8Array | ArrayBuffer,
	password: string,
	options: DecryptOptions = {},
): Promise<boolean> {
	try {
		const opened = openEncryptedPackage(input);
		if (!opened) return false;
		const info = parseEncryptionInfo(opened.encryptionInfo);
		const key = isStandard(info)
			? await verifyStandardPassword(info, password)
			: await verifyAgilePassword(info, password, options);
		return key !== null;
	} catch {
		return false;
	}
}
