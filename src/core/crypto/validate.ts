/**
 * Validation of untrusted EncryptionInfo parameters. The sizes in the stream drive buffer
 * slicing and the verifier comparison, so a hostile value (NaN, 0, a size that does not match
 * the digest) must be rejected rather than shortening or skipping the comparison.
 */

import { normalizeDigestAlgorithmName } from '../digest/algorithm-names';
import type { EncryptionInfo, StandardEncryptionInfo } from './types';

/** Digest lengths, in bytes, of the hashes [MS-OFFCRYPTO] 2.3.4.11 allows for agile encryption. */
const AGILE_DIGEST_LENGTH: Readonly<Record<string, number>> = {
	'SHA-1': 20,
	'SHA-256': 32,
	'SHA-384': 48,
	'SHA-512': 64,
};

/** The digest length of an agile `hashAlgorithm` name, or undefined when it is not allowed. */
export function agileDigestLength(hashAlgorithm: string): number | undefined {
	const name = normalizeDigestAlgorithmName(hashAlgorithm);
	return name === undefined ? undefined : AGILE_DIGEST_LENGTH[name];
}

const KEY_BITS = new Set([128, 192, 256]);

type Block = EncryptionInfo['keyData'] | EncryptionInfo['passwordKeyEncryptor'];

function assertBlock(block: Block, what: string): void {
	if (block.cipherAlgorithm !== 'AES') {
		throw new Error(`Unsupported ${what} cipherAlgorithm: ${block.cipherAlgorithm}.`);
	}
	if (block.cipherChaining !== 'ChainingModeCBC') {
		throw new Error(`Unsupported ${what} cipherChaining: ${block.cipherChaining}.`);
	}
	const digestLength = agileDigestLength(block.hashAlgorithm);
	if (digestLength === undefined) {
		throw new Error(`Unsupported ${what} hashAlgorithm: ${block.hashAlgorithm}.`);
	}
	if (block.hashSize !== digestLength) {
		throw new Error(
			`Invalid ${what} hashSize ${block.hashSize}: ${block.hashAlgorithm} produces ${digestLength} bytes.`,
		);
	}
	if (!KEY_BITS.has(block.keyBits)) throw new Error(`Invalid ${what} keyBits: ${block.keyBits}.`);
	if (block.blockSize !== 16) throw new Error(`Invalid ${what} blockSize: ${block.blockSize}.`);
	if (!Number.isInteger(block.saltSize) || block.saltSize <= 0 || block.saltSize > 64) {
		throw new Error(`Invalid ${what} saltSize: ${block.saltSize}.`);
	}
	if (block.saltValue.length !== block.saltSize) {
		throw new Error(`The ${what} saltValue does not match its saltSize.`);
	}
}

/**
 * Rejects agile EncryptionInfo parameters that are not valid for the scheme.
 *
 * @throws Error naming the offending field.
 */
export function assertValidAgileEncryptionInfo(info: EncryptionInfo): void {
	assertBlock(info.keyData, 'keyData');
	const pke = info.passwordKeyEncryptor;
	assertBlock(pke, 'password key encryptor');
	if (!Number.isInteger(pke.spinCount) || pke.spinCount < 0) {
		throw new Error(`Invalid spinCount: ${pke.spinCount}.`);
	}
	if (pke.saltSize > pke.hashSize) {
		// The verifier hash input is `saltSize` bytes of a hash-sized plaintext.
		throw new Error(`Invalid password key encryptor saltSize: ${pke.saltSize}.`);
	}
}

/**
 * Rejects Standard EncryptionInfo parameters that are not valid for the AES scheme (the verifier
 * is a SHA-1 hash, 20 bytes).
 *
 * @throws Error naming the offending field.
 */
export function assertValidStandardEncryptionInfo(info: StandardEncryptionInfo): void {
	if (!KEY_BITS.has(info.header.keySize)) {
		throw new Error(`Invalid Standard encryption keySize: ${info.header.keySize}.`);
	}
	if (info.verifier.salt.length !== 16 || info.verifier.saltSize !== 16) {
		throw new Error(`Invalid Standard encryption saltSize: ${info.verifier.saltSize}.`);
	}
	if (info.verifier.verifierHashSize !== 20) {
		throw new Error(
			`Invalid Standard encryption verifierHashSize: ${info.verifier.verifierHashSize}.`,
		);
	}
	if (
		info.verifier.encryptedVerifier.length !== 16 ||
		info.verifier.encryptedVerifierHash.length < 20
	) {
		throw new Error('The Standard encryption verifier is truncated.');
	}
}
