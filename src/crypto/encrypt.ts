/**
 * OOXML package encryption: the agile scheme (Office 2010+, the default) and the ECMA-376
 * Standard scheme (Office 2007), each producing a compound file with `EncryptionInfo` and
 * `EncryptedPackage` streams.
 *
 * Moved from `src/pptx/core/utils/ooxml-crypto-encrypt.ts` and the agile half of
 * `ooxml-crypto.ts` (see PROVENANCE.md).
 */

import { buildOle2 } from '@christophervr/ole2/ole2-parser-write';
import {
	aesCbcEncryptNoPad,
	getCrypto,
	hash,
	hmac,
	uint32LE,
} from '@christophervr/ole2/utils/ooxml-crypto-primitives';
import { aesEcbEncrypt } from './aes-ecb.js';
import {
	STANDARD_ALG_ID_HASH_SHA1,
	buildAgileEncryptionInfoXml,
	buildEncryptionInfoStream,
	buildStandardEncryptionInfoStream,
} from './encryption-info-write.js';
import {
	BLOCK_KEYS,
	computeAgileKeyBase,
	deriveAgileKeyFromBase,
	deriveStandardKey,
	generateIV,
} from './key-derivation.js';
import type { EncryptionInfo, EncryptionOptions } from './types.js';

const SEGMENT_SIZE = 4096;

/** A compound file holding the two encryption streams. */
function container(encryptionInfo: Uint8Array, encryptedPackage: Uint8Array): Uint8Array {
	const streams = new Map<string, Uint8Array>();
	streams.set('EncryptionInfo', encryptionInfo);
	streams.set('EncryptedPackage', encryptedPackage);
	return new Uint8Array(buildOle2(streams));
}

/** `size` random bytes. */
function randomBytes(size: number): Uint8Array {
	const bytes = new Uint8Array(size);
	getCrypto().getRandomValues(bytes);
	return bytes;
}

/** `data` zero-padded up to a multiple of `blockSize`. */
function padTo(data: Uint8Array, blockSize: number): Uint8Array {
	const padded = new Uint8Array(Math.ceil(data.length / blockSize) * blockSize);
	padded.set(data);
	return padded;
}

/** The 8-byte little-endian size prefix followed by `body`. */
function withSizePrefix(size: number, body: Uint8Array): Uint8Array {
	const out = new Uint8Array(8 + body.length);
	const view = new DataView(out.buffer, 0, 8);
	view.setUint32(0, size % 0x100000000, true);
	view.setUint32(4, Math.floor(size / 0x100000000), true);
	out.set(body, 8);
	return out;
}

/**
 * Encrypt a package with the agile scheme: each 4096-byte segment is encrypted separately with
 * an IV derived from the segment index. Returns the bytes with the 8-byte size prefix.
 */
export async function encryptAgilePackage(
	packageData: Uint8Array,
	key: Uint8Array,
	info: EncryptionInfo,
): Promise<Uint8Array> {
	const { hashAlgorithm, saltValue, blockSize } = info.keyData;
	const padded = padTo(packageData, SEGMENT_SIZE);
	const encrypted = withSizePrefix(packageData.length, padded);
	for (let start = 0; start < padded.length; start += SEGMENT_SIZE) {
		const iv = await generateIV(
			hashAlgorithm,
			saltValue,
			uint32LE(start / SEGMENT_SIZE),
			blockSize,
		);
		const segment = padded.subarray(start, start + SEGMENT_SIZE);
		encrypted.set(await aesCbcEncryptNoPad(key, iv, segment), 8 + start);
	}
	return encrypted;
}

/**
 * Encrypt a package with the ECMA-376 Standard scheme: AES-ECB under the password-derived key
 * ([MS-OFFCRYPTO] 2.3.4.15). Returns the bytes with the 8-byte size prefix.
 */
export async function encryptStandardPackage(
	packageData: Uint8Array,
	key: Uint8Array,
): Promise<Uint8Array> {
	const encryptedData = await aesEcbEncrypt(key, packageData);
	return withSizePrefix(packageData.length, encryptedData);
}

/** The agile scheme (AES-CBC, SHA-512, 100,000 spins by default) as a compound file. */
export async function encryptAgile(
	packageData: Uint8Array,
	password: string,
	options: EncryptionOptions = {},
): Promise<Uint8Array> {
	const keyBits = options.algorithm === 'AES128' ? 128 : 256;
	const hashAlgorithm = 'SHA-512';
	const hashSize = 64;
	const blockSize = 16;
	const spinCount = options.spinCount ?? 100000;
	const keyDataSalt = randomBytes(16);
	const pkeSalt = randomBytes(16);
	const documentKey = randomBytes(keyBits / 8);

	// Password verifier: a random input, its hash, both encrypted under password-derived keys.
	const verifierHashInput = randomBytes(16);
	const verifierHash = await hash(hashAlgorithm, verifierHashInput);
	const base = await computeAgileKeyBase(password, pkeSalt, spinCount, hashAlgorithm);
	const encryptUnder = async (block: Uint8Array, data: Uint8Array): Promise<Uint8Array> => {
		const key = await deriveAgileKeyFromBase(base, hashAlgorithm, block, keyBits, hashSize);
		return aesCbcEncryptNoPad(key, pkeSalt, padTo(data, blockSize));
	};
	const passwordKeyEncryptor: EncryptionInfo['passwordKeyEncryptor'] = {
		saltSize: 16,
		blockSize,
		keyBits,
		hashSize,
		cipherAlgorithm: 'AES',
		cipherChaining: 'ChainingModeCBC',
		hashAlgorithm,
		saltValue: pkeSalt,
		spinCount,
		encryptedVerifierHashInput: await encryptUnder(BLOCK_KEYS.verifierHashInput, verifierHashInput),
		encryptedVerifierHashValue: await encryptUnder(BLOCK_KEYS.verifierHashValue, verifierHash),
		encryptedKeyValue: await encryptUnder(BLOCK_KEYS.encryptedKeyValue, documentKey),
	};
	const keyData: EncryptionInfo['keyData'] = {
		saltSize: 16,
		blockSize,
		keyBits,
		hashSize,
		cipherAlgorithm: 'AES',
		cipherChaining: 'ChainingModeCBC',
		hashAlgorithm,
		saltValue: keyDataSalt,
	};
	const info: EncryptionInfo = {
		version: { major: 4, minor: 4 },
		isAgile: true,
		keyData,
		passwordKeyEncryptor,
	};
	const encryptedPackage = await encryptAgilePackage(packageData, documentKey, info);

	// Data integrity: an HMAC over the complete EncryptedPackage stream (size prefix included),
	// its random key and value encrypted under the document key.
	const hmacKey = randomBytes(hashSize);
	const hmacValue = await hmac(hashAlgorithm, hmacKey, encryptedPackage);
	const encryptIntegrity = async (block: Uint8Array, data: Uint8Array): Promise<Uint8Array> => {
		const iv = await generateIV(hashAlgorithm, keyDataSalt, block, blockSize);
		return aesCbcEncryptNoPad(documentKey, iv, padTo(data, blockSize));
	};
	const dataIntegrity = {
		encryptedHmacKey: await encryptIntegrity(BLOCK_KEYS.dataIntegrityHmacKey, hmacKey),
		encryptedHmacValue: await encryptIntegrity(BLOCK_KEYS.dataIntegrityHmacValue, hmacValue),
	};
	const xml = buildAgileEncryptionInfoXml(keyData, passwordKeyEncryptor, dataIntegrity);
	return container(buildEncryptionInfoStream(xml), encryptedPackage);
}

/**
 * The ECMA-376 Standard scheme (Office 2007-compatible) as a compound file. There is no
 * separate document key: the password-derived key encrypts the whole package (AES-ECB).
 */
export async function encryptStandard(
	packageData: Uint8Array,
	password: string,
	options: EncryptionOptions = {},
): Promise<Uint8Array> {
	const keyBits = options.algorithm === 'AES128' ? 128 : 256;
	const salt = randomBytes(16);
	const key = await deriveStandardKey(password, salt, keyBits, STANDARD_ALG_ID_HASH_SHA1);

	// A random 16-byte verifier plus its SHA-1 hash, padded to 32 bytes.
	const verifier = randomBytes(16);
	const paddedVerifierHash = new Uint8Array(32);
	paddedVerifierHash.set((await hash('SHA-1', verifier)).subarray(0, 20));
	const encryptionInfo = buildStandardEncryptionInfoStream(
		keyBits,
		salt,
		await aesEcbEncrypt(key, verifier),
		await aesEcbEncrypt(key, paddedVerifierHash),
	);
	return container(encryptionInfo, await encryptStandardPackage(packageData, key));
}
