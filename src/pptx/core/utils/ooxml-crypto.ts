// Compatibility entry point. ECMA-376 package encryption moved to the shared crypto area
// (src/crypto, `ooxml-core/crypto`): decryptOoxmlPackage, encryptOoxmlPackage and
// verifyOoxmlPackagePassword. The PowerPoint names below keep their ArrayBuffer signatures.
import {
	decryptOoxmlPackage,
	encryptOoxmlPackage,
	verifyOoxmlPackagePassword,
} from '../../../crypto/index.js';
import type { EncryptionOptions } from '../../../crypto/types.js';

export type {
	EncryptionAlgorithm,
	EncryptionInfo,
	StandardEncryptionInfo,
	EncryptionOptions,
	EncryptionScheme,
} from '../../../crypto/types.js';
export { IncorrectPasswordError, DataIntegrityError } from '../../../crypto/errors.js';
export {
	parseEncryptionInfo as _parseEncryptionInfo,
	deriveAgileKey as _deriveAgileKey,
} from '../../../crypto/index.js';
export {
	base64Decode as _base64Decode,
	base64Encode as _base64Encode,
	encodePasswordUtf16LE as _encodePasswordUtf16LE,
	concatArrays as _concatArrays,
	hash as _hash,
} from './ooxml-crypto-primitives';

/** Decrypt a password-protected PPTX file (see `decryptOoxmlPackage`). */
export async function decryptPptx(
	encryptedBuffer: ArrayBuffer,
	password: string,
): Promise<ArrayBuffer> {
	return (await decryptOoxmlPackage(encryptedBuffer, password)).buffer as ArrayBuffer;
}

/** Encrypt a PPTX file with a password (see `encryptOoxmlPackage`). */
export async function encryptPptx(
	pptxBuffer: ArrayBuffer,
	password: string,
	options?: EncryptionOptions,
): Promise<ArrayBuffer> {
	return (await encryptOoxmlPackage(pptxBuffer, password, options)).buffer as ArrayBuffer;
}

/** Whether a password opens an encrypted file (see `verifyOoxmlPackagePassword`). */
export function verifyPassword(encryptedBuffer: ArrayBuffer, password: string): Promise<boolean> {
	return verifyOoxmlPackagePassword(encryptedBuffer, password);
}
