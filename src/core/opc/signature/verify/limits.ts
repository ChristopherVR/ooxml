/**
 * The most work the verifier does for one signature. Each bound sits far above what an Office
 * signature needs, so only a crafted signature reaches one, and that signature is reported as not
 * checked rather than verified or invalid.
 */
export const SIGNATURE_LIMITS = {
	/** Largest signature part read: a signature holds digests and a few certificates. */
	maxSignaturePartBytes: 16 * 1024 * 1024,
	/** Smallest RSA modulus checked, in bits. */
	minRsaModulusBits: 1024,
	/** Largest RSA modulus checked, in bits. The cost of one check grows with its square. */
	maxRsaModulusBits: 16_384,
	/** Largest RSA public exponent checked. The cost of one check grows with its length. */
	maxRsaExponent: 2n ** 32n,
	/** Most `KeyInfo` certificates tried against `SignatureValue`. */
	maxCertificatesTried: 4,
	/** Most references read, `Manifest` references included: one per part of the largest package. */
	maxReferences: 10_000,
	/** Most part bytes read and digested: every part of the largest package once. */
	maxDigestedBytes: 100 * 1024 * 1024,
} as const;
