/**
 * OOXML encryption key derivation ([MS-OFFCRYPTO] 2.3.4.7 and 2.3.4.11).
 *
 * The password spin (100,000 rounds for agile, 50,000 for Standard) runs on the synchronous
 * digests of the shared `digest` area when the hash is one it computes, so a key is derived
 * without awaiting Web Crypto once per round; other hashes fall back to Web Crypto.
 *
 * Moved from `src/pptx/core/utils/ooxml-crypto-key-derivation.ts` (see PROVENANCE.md).
 */

import {
	concatArrays,
	encodePasswordUtf16LE,
	hash,
	uint32LE,
} from '@christophervr/ole2/utils/ooxml-crypto-primitives';
import { digestFunction } from '../digest/digest.js';

/** Well-known block keys for agile encryption as defined in [MS-OFFCRYPTO]. */
export const BLOCK_KEYS = {
	verifierHashInput: new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
	verifierHashValue: new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
	encryptedKeyValue: new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
	dataIntegrityHmacKey: new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
	dataIntegrityHmacValue: new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]),
} as const;

/** The largest spin count honoured; Office writes 100,000, far more is a hostile file. */
export const MAX_SPIN_COUNT = 10_000_000;

/** H0 = H(salt + password), then `spinCount` rounds of Hn = H(iterator + Hn-1). */
async function spinPassword(
	password: string,
	salt: Uint8Array,
	spinCount: number,
	hashAlgorithm: string,
): Promise<Uint8Array> {
	if (!Number.isInteger(spinCount) || spinCount < 0 || spinCount > MAX_SPIN_COUNT) {
		throw new Error(`Unsupported encryption spin count: ${spinCount}.`);
	}
	const passwordBytes = encodePasswordUtf16LE(password);
	const digest = digestFunction(hashAlgorithm);
	if (digest) {
		let h = digest(concatArrays(salt, passwordBytes));
		const buffer = new Uint8Array(4 + h.length);
		const view = new DataView(buffer.buffer);
		for (let i = 0; i < spinCount; i++) {
			view.setUint32(0, i, true);
			buffer.set(h, 4);
			h = digest(buffer);
		}
		return h;
	}
	let h = await hash(hashAlgorithm, concatArrays(salt, passwordBytes));
	for (let i = 0; i < spinCount; i++) {
		h = await hash(hashAlgorithm, concatArrays(uint32LE(i), h));
	}
	return h;
}

/** X1 = H(h ^ 0x36 padded to 64), X2 = H(h ^ 0x5C padded to 64), key = (X1 + X2) truncated. */
async function extendKey(
	h: Uint8Array,
	hashAlgorithm: string,
	keyLength: number,
): Promise<Uint8Array> {
	const pad = async (value: number): Promise<Uint8Array> => {
		const input = new Uint8Array(64).fill(value);
		for (let i = 0; i < h.length && i < 64; i++) input[i] = h[i]! ^ value;
		return hash(hashAlgorithm, input);
	};
	const x3 = concatArrays(await pad(0x36), await pad(0x5c));
	return x3.subarray(0, keyLength);
}

/**
 * The spun password hash for agile key derivation (the expensive part). Derive several block
 * keys from one base with {@link deriveAgileKeyFromBase} instead of spinning again for each.
 */
export function computeAgileKeyBase(
	password: string,
	salt: Uint8Array,
	spinCount: number,
	hashAlgorithm: string,
): Promise<Uint8Array> {
	return spinPassword(password, salt, spinCount, hashAlgorithm);
}

/** Finishes agile key derivation for one block key: Hfinal = H(base + blockKey), sized. */
export async function deriveAgileKeyFromBase(
	base: Uint8Array,
	hashAlgorithm: string,
	blockKey: Uint8Array,
	keyBits: number,
	hashSize: number,
): Promise<Uint8Array> {
	const h = await hash(hashAlgorithm, concatArrays(base, blockKey));
	const keyLength = keyBits / 8;
	if (hashSize >= keyLength) return h.subarray(0, keyLength);
	return extendKey(h, hashAlgorithm, keyLength);
}

/**
 * Derive an encryption key from a password using the OOXML agile encryption key derivation
 * algorithm.
 *
 * @param password - User's password.
 * @param salt - Salt from EncryptionInfo.
 * @param spinCount - Number of hash iterations.
 * @param hashAlgorithm - Hash algorithm name (e.g. "SHA-512").
 * @param blockKey - Block key for deriving specific sub-keys.
 * @param keyBits - Desired key length in bits.
 * @param hashSize - Hash output size in bytes.
 * @returns Derived key of keyBits/8 bytes.
 */
export async function deriveAgileKey(
	password: string,
	salt: Uint8Array,
	spinCount: number,
	hashAlgorithm: string,
	blockKey: Uint8Array,
	keyBits: number,
	hashSize: number,
): Promise<Uint8Array> {
	const base = await computeAgileKeyBase(password, salt, spinCount, hashAlgorithm);
	return deriveAgileKeyFromBase(base, hashAlgorithm, blockKey, keyBits, hashSize);
}

/**
 * Compute the expensive, password/salt-dependent part of Standard Encryption key derivation:
 * `H50000`, the result of hashing salt+password and then iterating 50,000 times.
 *
 * Only OOXML Standard Encryption spins. The legacy `.ppt` "RC4 CryptoAPI" scheme has no spin
 * rounds and lives in the pptx area (`ppt/rc4-cryptoapi-key.ts`).
 */
export function computeStandardKeyBase(password: string, salt: Uint8Array): Promise<Uint8Array> {
	return spinPassword(password, salt, 50000, 'SHA-1');
}

/**
 * Finish Standard Encryption key derivation for one block, given the expensive base from
 * {@link computeStandardKeyBase}: `Hfinal = H(H50000 + blockKey)`, then always the X1/X2/X3
 * extension, of which the key is the first `keySize` bits ([MS-OFFCRYPTO] 2.3.4.7). Unlike the
 * binary-document RC4 CryptoAPI derivation (2.3.5.2), a 128-bit key is not a plain truncation of
 * Hfinal.
 */
export async function deriveStandardKeyFromBase(
	base: Uint8Array,
	keySize: number,
	blockNumber = 0,
): Promise<Uint8Array> {
	const h = await hash('SHA-1', concatArrays(base, uint32LE(blockNumber)));
	return extendKey(h, 'SHA-1', keySize / 8);
}

/**
 * Derive the encryption key for OOXML Standard Encryption (Office 2007), [MS-OFFCRYPTO]
 * 2.3.4.7 (50,000 spin rounds).
 *
 * @param password - User's password.
 * @param salt - Salt from the verifier.
 * @param keySize - Key size in bits (e.g. 128).
 * @param _algIdHash - Algorithm ID for hashing (always SHA-1 in practice).
 * @param blockNumber - Block number (0 for Standard Encryption).
 */
export async function deriveStandardKey(
	password: string,
	salt: Uint8Array,
	keySize: number,
	_algIdHash: number,
	blockNumber = 0,
): Promise<Uint8Array> {
	const base = await computeStandardKeyBase(password, salt);
	return deriveStandardKeyFromBase(base, keySize, blockNumber);
}

/**
 * Generate an IV for agile encryption: H(salt + blockKey), truncated to `blockSize` or padded
 * with 0x36.
 */
export async function generateIV(
	hashAlgorithm: string,
	salt: Uint8Array,
	blockKey: Uint8Array,
	blockSize: number,
): Promise<Uint8Array> {
	const h = await hash(hashAlgorithm, concatArrays(salt, blockKey));
	if (h.length >= blockSize) return h.subarray(0, blockSize);
	const padded = new Uint8Array(blockSize).fill(0x36);
	padded.set(h);
	return padded;
}
