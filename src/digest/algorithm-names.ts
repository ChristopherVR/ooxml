/**
 * Canonical spellings of the hash algorithm names ECMA-376 lets a password verifier name.
 *
 * The `algorithmName` attributes (`sheetProtection`, `workbookProtection`, `p:modifyVerifier`,
 * `w:writeProtection`, ...) are free-form strings: ECMA-376 gives `"SHA-512"` and `"SHA512"`,
 * and hyphen-less, lower-case or underscored forms (`"sha1"`, `"sha_512"`, `"ripemd160"`) occur
 * in the wild. Normalising once means every caller agrees on one spelling.
 *
 * Ported from the pptx area's `src/pptx/core/utils/digests/algorithm-names.ts`.
 */

/** The hash algorithm names ECMA-376 password verifiers use and this module recognises. */
export type DigestAlgorithmName =
	| 'SHA-1'
	| 'SHA-256'
	| 'SHA-384'
	| 'SHA-512'
	| 'MD2'
	| 'MD4'
	| 'MD5'
	| 'RIPEMD-128'
	| 'RIPEMD-160'
	| 'WHIRLPOOL';

/** The algorithms {@link digestSync} computes (the others are recognised but not implemented). */
export type SupportedDigestName = 'SHA-1' | 'SHA-256' | 'SHA-384' | 'SHA-512';

/** Lookup key: the name upper-cased with hyphens, underscores and white space removed. */
const CANONICAL_BY_LOOKUP_KEY: Readonly<Record<string, DigestAlgorithmName>> = {
	SHA1: 'SHA-1',
	SHA256: 'SHA-256',
	SHA384: 'SHA-384',
	SHA512: 'SHA-512',
	MD2: 'MD2',
	MD4: 'MD4',
	MD5: 'MD5',
	RIPEMD128: 'RIPEMD-128',
	RIPEMD160: 'RIPEMD-160',
	WHIRLPOOL: 'WHIRLPOOL',
};

/**
 * Normalises `name` to a {@link DigestAlgorithmName}, or `undefined` when it names an algorithm
 * this module does not recognise.
 */
export function normalizeDigestAlgorithmName(name: string): DigestAlgorithmName | undefined {
	const key = name.toUpperCase().replace(/[-_\s]/g, '');
	return Object.hasOwn(CANONICAL_BY_LOOKUP_KEY, key) ? CANONICAL_BY_LOOKUP_KEY[key] : undefined;
}
