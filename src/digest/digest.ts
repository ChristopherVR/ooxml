import { normalizeDigestAlgorithmName, type SupportedDigestName } from './algorithm-names.js';
import { sha1 } from './sha1.js';
import { sha256 } from './sha256.js';
import { sha384, sha512 } from './sha512.js';

/** A synchronous hash function over bytes. */
export type DigestFunction = (data: Uint8Array) => Uint8Array;

const DIGESTS: Readonly<Record<SupportedDigestName, DigestFunction>> = {
	'SHA-1': sha1,
	'SHA-256': sha256,
	'SHA-384': sha384,
	'SHA-512': sha512,
};

/**
 * The synchronous digest function for `name` (any spelling {@link normalizeDigestAlgorithmName}
 * accepts), or `undefined` when it is not one this module computes.
 */
export function digestFunction(name: string): DigestFunction | undefined {
	const canonical = normalizeDigestAlgorithmName(name);
	return canonical && canonical in DIGESTS ? DIGESTS[canonical as SupportedDigestName] : undefined;
}

/**
 * The digest of `data` under `name` (SHA-1, SHA-256, SHA-384 or SHA-512, in any ECMA-376
 * spelling), or `undefined` for an algorithm this module does not compute.
 */
export function digestSync(name: string, data: Uint8Array): Uint8Array | undefined {
	return digestFunction(name)?.(data);
}
