import { normalizeDigestAlgorithmName, type SupportedDigestName } from './algorithm-names';
import { md2 } from './md2';
import { md4 } from './md4';
import { md5 } from './md5';
import { ripemd128 } from './ripemd128';
import { ripemd160 } from './ripemd160';
import { sha1 } from './sha1';
import { sha256 } from './sha256';
import { sha384, sha512 } from './sha512';
import { whirlpool } from './whirlpool';

/** A synchronous hash function over bytes. */
export type DigestFunction = (data: Uint8Array) => Uint8Array;

const DIGESTS: Readonly<Record<SupportedDigestName, DigestFunction>> = {
	'SHA-1': sha1,
	'SHA-256': sha256,
	'SHA-384': sha384,
	'SHA-512': sha512,
	MD2: md2,
	MD4: md4,
	MD5: md5,
	'RIPEMD-128': ripemd128,
	'RIPEMD-160': ripemd160,
	WHIRLPOOL: whirlpool,
};

/**
 * The synchronous digest function for `name` (any spelling {@link normalizeDigestAlgorithmName}
 * accepts), or `undefined` when it is not one this module computes.
 */
export function digestFunction(name: string): DigestFunction | undefined {
	const canonical = normalizeDigestAlgorithmName(name);
	return canonical && DIGESTS[canonical];
}

/**
 * The digest of `data` under `name` (SHA-1, SHA-256, SHA-384, SHA-512, MD2, MD4, MD5,
 * RIPEMD-128, RIPEMD-160 or WHIRLPOOL, in any ECMA-376 spelling), or `undefined` for an algorithm this module does not compute.
 */
export function digestSync(name: string, data: Uint8Array): Uint8Array | undefined {
	return digestFunction(name)?.(data);
}
