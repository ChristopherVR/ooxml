// Synchronous, pure-TypeScript message digests (SHA-1/256/384/512, MD2, MD4, MD5, RIPEMD-128,
// RIPEMD-160, WHIRLPOOL) and the ECMA-376 agile password hash shared by
// every format's protection elements. No Web Crypto: it works outside secure contexts (http://)
// and runs the 100,000-round Office password spin without awaiting each round.
export {
	normalizeDigestAlgorithmName,
	type DigestAlgorithmName,
	type SupportedDigestName,
} from './algorithm-names';
export { decodeBase64, encodeBase64 } from './base64';
export { digestFunction, digestSync, type DigestFunction } from './digest';
export { md2 } from './md2';
export { md4 } from './md4';
export { md5 } from './md5';
export {
	MAX_SPIN_COUNT,
	hashPassword,
	spinPasswordHash,
	verifyPasswordHash,
	type PasswordHash,
} from './password-hash';
export { ripemd128 } from './ripemd128';
export { ripemd160 } from './ripemd160';
export { sha1 } from './sha1';
export { sha256 } from './sha256';
export { sha384, sha512 } from './sha512';
export { whirlpool } from './whirlpool';
