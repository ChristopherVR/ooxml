// OOXML package encryption ([MS-OFFCRYPTO]), shared by every format: a password-to-open .xlsx,
// .docx or .pptx is a compound file (CFB, from ole2, inlined at build time) holding an
// `EncryptionInfo` stream and the AES-encrypted zip in `EncryptedPackage`. Agile (Office 2010+)
// and Standard (Office 2007) are decrypted; agile is written by default.
//
// Moved from the pptx area (`src/pptx/core/utils/ooxml-crypto*.ts`, `encryption-detection.ts`;
// see PROVENANCE.md). Not re-exported from the root entry: it pulls the CFB codec in.
import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';
import {
	decryptAgilePackage,
	decryptStandardPackage,
	verifyAgileDataIntegrity,
	verifyAgilePassword,
	verifyStandardPassword,
} from './decrypt.js';
import { openEncryptedPackage, toArrayBuffer, toBytes } from './detect.js';
import { encryptAgile, encryptStandard } from './encrypt.js';
import { parseEncryptionInfo } from './encryption-info.js';
import { IncorrectPasswordError } from './errors.js';
import type { EncryptionInfo, EncryptionOptions, StandardEncryptionInfo } from './types.js';

export {
	DataIntegrityError,
	IncorrectPasswordError,
	PasswordRequiredError,
	isOoxmlCryptoError,
	type OoxmlCryptoErrorCode,
} from './errors.js';
export {
	EncryptedFileError,
	detectFileFormat,
	isCompoundFile,
	isEncryptedOoxmlPackage,
	openEncryptedPackage,
	type EncryptedPackageStreams,
	type FileFormatDetection,
} from './detect.js';
export type {
	EncryptionAlgorithm,
	EncryptionInfo,
	EncryptionOptions,
	EncryptionScheme,
	StandardEncryptionInfo,
} from './types.js';
export { RC4_ALG_ID, parseEncryptionInfo } from './encryption-info.js';
export {
	BLOCK_KEYS,
	computeStandardKeyBase,
	deriveAgileKey,
	deriveStandardKey,
	deriveStandardKeyFromBase,
	generateIV,
} from './key-derivation.js';
export {
	decryptAgilePackage,
	decryptStandardPackage,
	verifyAgileDataIntegrity,
	verifyAgilePassword,
	verifyStandardPassword,
} from './decrypt.js';
export { encryptAgilePackage, encryptStandardPackage } from './encrypt.js';
export {
	buildAgileEncryptionInfoXml,
	buildEncryptionInfoStream,
	buildStandardEncryptionInfoStream,
} from './encryption-info-write.js';

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
 * @throws DataIntegrityError (`code: 'data-integrity'`) when the agile HMAC does not match.
 * @throws Error when the bytes are not an encrypted package or the scheme is unsupported.
 */
export async function decryptOoxmlPackage(
	input: Uint8Array | ArrayBuffer,
	password: string,
): Promise<Uint8Array> {
	const { info: infoStream, data } = streams(toBytes(input));
	const info = parseEncryptionInfo(infoStream);
	if (isStandard(info)) {
		const key = await verifyStandardPassword(info, password);
		if (!key) throw new IncorrectPasswordError();
		return new Uint8Array(await decryptStandardPackage(data, key));
	}
	const key = await verifyAgilePassword(info, password);
	if (!key) throw new IncorrectPasswordError();
	await verifyAgileDataIntegrity(info, key, data);
	return new Uint8Array(await decryptAgilePackage(data, key, info));
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
): Promise<boolean> {
	try {
		const opened = openEncryptedPackage(input);
		if (!opened) return false;
		const info = parseEncryptionInfo(opened.encryptionInfo);
		const key = isStandard(info)
			? await verifyStandardPassword(info, password)
			: await verifyAgilePassword(info, password);
		return key !== null;
	} catch {
		return false;
	}
}
