/**
 * OOXML package decryption: password verification and package decryption for the agile
 * (Office 2010+) and Standard (Office 2007) encryption schemes.
 *
 * Moved from `src/core/pptx/core/utils/ooxml-crypto-decrypt.ts` (see PROVENANCE.md).
 */

import {
	aesCbcDecryptRaw,
	hash,
	hmac,
	uint32LE,
} from '@christophervr/ole2/utils/ooxml-crypto-primitives';
import type { SpinLimits } from '../digest/password-hash';
import { timingSafeEqual } from '../digest/timing-safe';
import { aesEcbDecrypt } from './aes-ecb';
import { DataIntegrityError } from './errors';
import {
	BLOCK_KEYS,
	computeAgileKeyBase,
	deriveAgileKeyFromBase,
	deriveStandardKey,
	generateIV,
} from './key-derivation';
import type { EncryptionInfo, StandardEncryptionInfo } from './types';
import { assertValidAgileEncryptionInfo, assertValidStandardEncryptionInfo } from './validate';

/** Options of agile password verification and package decryption. */
export interface DecryptOptions extends SpinLimits {
	/**
	 * Accept an agile package that has no `<dataIntegrity>` element (default false). The HMAC is
	 * the only protection of the ciphertext against tampering, so a missing element is treated as
	 * a downgrade attack unless the caller opts out (for example for a legacy writer).
	 */
	allowMissingIntegrity?: boolean;
}

// ---------------------------------------------------------------------------
// Agile Password Verification
// ---------------------------------------------------------------------------

/**
 * Verify the password against agile encryption info and return the
 * decryption key if valid.
 *
 * @param info - Parsed agile encryption info.
 * @param password - The password to verify.
 * @param limits - The spin-count limit (default 1,000,000; the spec ceiling is 10,000,000).
 * @returns The document encryption key, or null if the password is wrong.
 * @throws Error when the EncryptionInfo parameters are invalid or the spin count is too large.
 */
export async function verifyAgilePassword(
	info: EncryptionInfo,
	password: string,
	limits: SpinLimits = {},
): Promise<Uint8Array | null> {
	assertValidAgileEncryptionInfo(info);
	const pke = info.passwordKeyEncryptor;
	// The spin is the expensive part and depends only on the password and salt: run it once.
	const base = await computeAgileKeyBase(
		password,
		pke.saltValue,
		pke.spinCount,
		pke.hashAlgorithm,
		limits,
	);
	const blockKey = (block: Uint8Array): Promise<Uint8Array> =>
		deriveAgileKeyFromBase(base, pke.hashAlgorithm, block, pke.keyBits, pke.hashSize);

	const verifierHashInput = await aesCbcDecryptRaw(
		await blockKey(BLOCK_KEYS.verifierHashInput),
		pke.saltValue,
		pke.encryptedVerifierHashInput,
	);
	const verifierHashValue = await aesCbcDecryptRaw(
		await blockKey(BLOCK_KEYS.verifierHashValue),
		pke.saltValue,
		pke.encryptedVerifierHashValue,
	);

	// Hash the decrypted verifier input and compare up to the hash size.
	const computedHash = await hash(pke.hashAlgorithm, verifierHashInput.subarray(0, pke.saltSize));
	// hashSize was validated against the digest length, so the whole digest is compared.
	if (computedHash.length !== pke.hashSize || verifierHashValue.length < pke.hashSize) return null;
	if (!timingSafeEqual(verifierHashValue.subarray(0, pke.hashSize), computedHash)) return null;

	// Password verified: decrypt the document encryption key.
	const decryptedKey = await aesCbcDecryptRaw(
		await blockKey(BLOCK_KEYS.encryptedKeyValue),
		pke.saltValue,
		pke.encryptedKeyValue,
	);
	return decryptedKey.subarray(0, info.keyData.keyBits / 8);
}

// ---------------------------------------------------------------------------
// Standard Password Verification
// ---------------------------------------------------------------------------

/**
 * Verify the password against standard encryption info and return the
 * decryption key if valid.
 *
 * @param info - Parsed standard encryption info.
 * @param password - The password to verify.
 * @returns The encryption key, or null if the password is wrong.
 */
export async function verifyStandardPassword(
	info: StandardEncryptionInfo,
	password: string,
): Promise<Uint8Array | null> {
	assertValidStandardEncryptionInfo(info);
	const key = await deriveStandardKey(
		password,
		info.verifier.salt,
		info.header.keySize,
		info.header.algIdHash,
	);

	// Decrypt the encrypted verifier
	// The Standard scheme encrypts the verifier and its hash with AES-ECB ([MS-OFFCRYPTO] 2.3.4.8).
	const decryptedVerifier = await aesEcbDecrypt(key, info.verifier.encryptedVerifier);

	// Decrypt the encrypted verifier hash
	const decryptedHash = await aesEcbDecrypt(key, info.verifier.encryptedVerifierHash);

	// Hash the decrypted verifier
	const computedHash = await hash('SHA-1', decryptedVerifier);

	// The verifier hash is SHA-1 (20 bytes, enforced by the validation above): compare all of it.
	const match = timingSafeEqual(decryptedHash.subarray(0, 20), computedHash);

	return match ? key : null;
}

// ---------------------------------------------------------------------------
// Data Integrity Verification
// ---------------------------------------------------------------------------

/**
 * Verify the data integrity HMAC of an agile-encrypted package.
 *
 * [MS-OFFCRYPTO] 2.3.7.1 -- The data integrity is verified by:
 * 1. Decrypting the HMAC key from dataIntegrity.encryptedHmacKey
 * 2. Decrypting the HMAC value from dataIntegrity.encryptedHmacValue
 * 3. Computing HMAC of the encrypted package data (after the 8-byte size prefix)
 * 4. Comparing the computed HMAC with the decrypted HMAC value
 *
 * @param info - Parsed agile encryption info.
 * @param key - The document encryption key.
 * @param encryptedPackage - Raw bytes of the EncryptedPackage stream.
 * @param options - `allowMissingIntegrity` accepts a package without a dataIntegrity element.
 * @throws DataIntegrityError if the data integrity check fails, or the element is missing and
 *   `allowMissingIntegrity` is not set (an attacker could strip it to flip ciphertext bits).
 */
export async function verifyAgileDataIntegrity(
	info: EncryptionInfo,
	key: Uint8Array,
	encryptedPackage: Uint8Array,
	options: Pick<DecryptOptions, 'allowMissingIntegrity'> = {},
): Promise<void> {
	const keyData = info.keyData;

	if (!info.dataIntegrity) {
		if (options.allowMissingIntegrity) return;
		throw new DataIntegrityError(
			'The encrypted package has no data integrity record, so it cannot be checked for tampering.',
		);
	}

	// Decrypt the HMAC key
	const hmacKeyIV = await generateIV(
		keyData.hashAlgorithm,
		keyData.saltValue,
		BLOCK_KEYS.dataIntegrityHmacKey,
		keyData.blockSize,
	);
	const decryptedHmacKey = await aesCbcDecryptRaw(
		key,
		hmacKeyIV,
		info.dataIntegrity.encryptedHmacKey,
	);
	// Truncate to hash size
	const hmacKey = decryptedHmacKey.subarray(0, keyData.hashSize);

	// Decrypt the HMAC value
	const hmacValueIV = await generateIV(
		keyData.hashAlgorithm,
		keyData.saltValue,
		BLOCK_KEYS.dataIntegrityHmacValue,
		keyData.blockSize,
	);
	const decryptedHmacValue = await aesCbcDecryptRaw(
		key,
		hmacValueIV,
		info.dataIntegrity.encryptedHmacValue,
	);
	const expectedHmac = decryptedHmacValue.subarray(0, keyData.hashSize);

	// The HMAC covers the complete EncryptedPackage stream, including its size prefix.
	const computedHmac = await hmac(keyData.hashAlgorithm, hmacKey, encryptedPackage);

	// Compare HMACs (a length mismatch fails)
	const match = timingSafeEqual(computedHmac, expectedHmac);

	if (!match) {
		throw new DataIntegrityError(
			'Data integrity check failed. The encrypted file may be corrupted or tampered with.',
		);
	}
}

// ---------------------------------------------------------------------------
// Package Decryption
// ---------------------------------------------------------------------------

/** The 8-byte StreamSize prefix: the byte length of the original package. */
function readOriginalSize(encryptedPackage: Uint8Array, wide: boolean): number {
	if (encryptedPackage.length < 8) throw new Error('The EncryptedPackage stream is truncated.');
	const view = new DataView(encryptedPackage.buffer, encryptedPackage.byteOffset, 8);
	const high = view.getUint32(4, true);
	// Standard encryption only defines the low 32 bits.
	return view.getUint32(0, true) + (wide ? high * 0x100000000 : 0);
}

/**
 * Decrypt the EncryptedPackage stream using the agile encryption key.
 *
 * The encrypted package uses segment-based encryption:
 * each 4096-byte segment is encrypted separately with a unique IV.
 *
 * @param encryptedPackage - Raw bytes of the EncryptedPackage stream.
 * @param key - The document encryption key.
 * @param info - Parsed agile encryption info.
 * @returns The decrypted package as an ArrayBuffer.
 */
export async function decryptAgilePackage(
	encryptedPackage: Uint8Array,
	key: Uint8Array,
	info: EncryptionInfo,
): Promise<ArrayBuffer> {
	const keyData = info.keyData;

	const originalSize = readOriginalSize(encryptedPackage, true);

	const encryptedData = encryptedPackage.subarray(8);
	// Reject a declared size the ciphertext cannot hold before allocating the result from it.
	if (originalSize > encryptedData.length) {
		throw new Error('The declared package size exceeds the encrypted data.');
	}
	const segmentSize = 4096;
	const result = new Uint8Array(originalSize);
	let resultOffset = 0;

	const numSegments = Math.ceil(encryptedData.length / segmentSize);

	for (let segment = 0; segment < numSegments; segment++) {
		const segmentStart = segment * segmentSize;
		const segmentEnd = Math.min(segmentStart + segmentSize, encryptedData.length);
		const segmentData = encryptedData.subarray(segmentStart, segmentEnd);

		// Generate IV for this segment: H(salt + blockKey)
		const blockKeyBytes = uint32LE(segment);
		const segmentIV = await generateIV(
			keyData.hashAlgorithm,
			keyData.saltValue,
			blockKeyBytes,
			keyData.blockSize,
		);

		const decrypted = await aesCbcDecryptRaw(key, segmentIV, segmentData);

		// Copy only what's needed (last segment might be smaller)
		const bytesToCopy = Math.min(decrypted.length, originalSize - resultOffset);
		result.set(decrypted.subarray(0, bytesToCopy), resultOffset);
		resultOffset += bytesToCopy;
	}

	return result.buffer;
}

/**
 * Decrypt the EncryptedPackage stream using standard encryption key.
 *
 * Standard encryption uses AES-ECB over the whole package ([MS-OFFCRYPTO] 2.3.4.15).
 *
 * @param encryptedPackage - Raw bytes of the EncryptedPackage stream.
 * @param key - The encryption key.
 * @returns The decrypted package as an ArrayBuffer.
 */
export async function decryptStandardPackage(
	encryptedPackage: Uint8Array,
	key: Uint8Array,
): Promise<ArrayBuffer> {
	const originalSize = readOriginalSize(encryptedPackage, false);

	const encryptedData = encryptedPackage.subarray(8);
	if (originalSize > encryptedData.length) {
		throw new Error('The declared package size exceeds the encrypted data.');
	}
	const decrypted = await aesEcbDecrypt(key, encryptedData);

	// Uint8Array#slice (unlike #subarray) copies into a freshly allocated
	// buffer sized to the given range, so the returned ArrayBuffer is
	// truncated to originalSize instead of exposing the padded backing
	// buffer of `decrypted`.
	return decrypted.slice(0, originalSize).buffer as ArrayBuffer;
}
