/**
 * Byte-level helpers for `p:modifyVerifier` password hashing: the base64
 * codec, salt generation, and the ECMA-376 19.2.1.22 / [MS-OFFCRYPTO] 2.3.7.1
 * iterated-hash derivation itself. Split out of `modify-verifier.ts` (which
 * owns the public API and the CAPI algorithm-name resolution) purely to keep
 * both files under the repo's per-file line budget; nothing here is useful on
 * its own outside that module.
 *
 * @module modify-verifier-codec
 */

import type { DigestAlgorithmName } from '../../../digest/algorithm-names';
import { digestFunction } from '../../../digest/digest';
import { isValidSpinCount, spinPasswordHash } from '../../../digest/password-hash';

/**
 * Decode base64 string to Uint8Array.
 *
 * Strips whitespace (spaces, tabs, newlines) first: `xsd:base64Binary`
 * permits embedded whitespace, and a hand-authored or line-wrapped
 * `p:modifyVerifier` attribute has been seen carrying it. `Buffer.from`
 * tolerates this in Node, but the browser's `atob` throws
 * `InvalidCharacterError` on it, so this normalisation has to happen here to
 * behave the same in both environments.
 */
export function base64Decode(str: string): Uint8Array {
	const normalized = str.replace(/\s+/g, '');
	if (typeof Buffer !== 'undefined') {
		const buf = Buffer.from(normalized, 'base64');
		return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
	}
	const binary = atob(normalized);
	const bytes = new Uint8Array(binary.length);
	for (let i = 0; i < binary.length; i++) {
		bytes[i] = binary.charCodeAt(i);
	}
	return bytes;
}

/** Encode Uint8Array to base64 string. */
export function base64Encode(bytes: Uint8Array): string {
	if (typeof Buffer !== 'undefined') {
		return Buffer.from(bytes).toString('base64');
	}
	let binary = '';
	for (let i = 0; i < bytes.length; i++) {
		binary += String.fromCharCode(bytes[i]!);
	}
	return btoa(binary);
}

/**
 * Random 16-byte salt (PowerPoint's own width).
 *
 * @throws Error where `crypto.getRandomValues` is unavailable: a predictable salt would defeat
 *   the verifier, so there is no fallback.
 */
export function randomSalt(): Uint8Array {
	const salt = new Uint8Array(16);
	if (typeof globalThis.crypto?.getRandomValues !== 'function') {
		throw new Error(
			'A secure random source (crypto.getRandomValues) is required to create a salt.',
		);
	}
	globalThis.crypto.getRandomValues(salt);
	return salt;
}

/**
 * The ECMA-376 19.2.1.22 / [MS-OFFCRYPTO] 2.3.7.1 iterated-hash derivation
 * shared by `verifyModifyPassword`, `createModifyVerifier` and
 * `createSaltlessModifyVerifierForTesting` (`./modify-verifier`):
 *
 * 1. H0 = H(salt + password_utf16le)
 * 2. For i = 0..spinCount-1: Hi = H(Hi-1 + i_le32)
 *
 * `salt` may be zero-length (a salt-less verifier): the formula is unchanged,
 * it simply prepends nothing to the password bytes. The iteration's byte
 * order (previous hash FIRST, iterator counter SECOND) was verified against
 * a `.pptx` produced by real PowerPoint (COM `Presentation.WritePassword`,
 * `p:modifyVerifier/@cryptAlgorithmSid="14"`, i.e. SHA-512): the reverse
 * order, used by an earlier version of this module, only ever passed its own
 * round-trip tests and could never verify an actual PowerPoint-authored
 * password. The derivation itself is the shared synchronous
 * `spinPasswordHash` (`ooxml-core/digest`), the same one xlsx protection uses.
 */
export async function iteratedHash(
	algorithm: DigestAlgorithmName,
	salt: Uint8Array,
	password: string,
	spinCount: number,
): Promise<Uint8Array> {
	// Every algorithm a verifier can name is computed synchronously by the shared `digest` area,
	// so the 100,000-round spin runs in one tight loop instead of awaiting each round.
	if (!isValidSpinCount(spinCount)) {
		throw new Error(`Unsupported verifier spin count: ${spinCount}.`);
	}
	const digest = digestFunction(algorithm);
	if (!digest) {
		throw new Error(`Unsupported digest algorithm: ${algorithm as string}`);
	}
	return spinPasswordHash(password, salt, spinCount, digest);
}
