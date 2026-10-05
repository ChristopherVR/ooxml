// Synchronous, pure-TypeScript message digests and the ECMA-376 agile password hash shared by
// every format's protection elements. No Web Crypto: it works outside secure contexts (http://)
// and runs the 100,000-round Office password spin without awaiting each round.
export {
	normalizeDigestAlgorithmName,
	type DigestAlgorithmName,
	type SupportedDigestName,
} from './algorithm-names.js';
export { decodeBase64, encodeBase64 } from './base64.js';
export { digestFunction, digestSync, type DigestFunction } from './digest.js';
export {
	MAX_SPIN_COUNT,
	hashPassword,
	verifyPasswordHash,
	type PasswordHash,
} from './password-hash.js';
export { sha1 } from './sha1.js';
export { sha256 } from './sha256.js';
export { sha384, sha512 } from './sha512.js';
