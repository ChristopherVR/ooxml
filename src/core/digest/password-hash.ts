import { decodeBase64, encodeBase64 } from './base64';
import { digestFunction, type DigestFunction } from './digest';
import { timingSafeEqual } from './timing-safe';

/**
 * An ECMA-376 agile password hash as Office stores it on protection elements (`sheetProtection`,
 * `workbookProtection`, `p:modifyVerifier`, ...): digest name, base64 salt and hash, and the
 * number of extra rounds (absent means none).
 */
export interface PasswordHash {
	algorithmName: string;
	hashValue: string;
	saltValue: string;
	spinCount?: number;
}

/**
 * The hard ceiling of `spinCount` the spec allows (10,000,000). A caller may raise its limit up
 * to this value with the `maxSpinCount` option; nothing above it is ever honoured.
 */
export const MAX_SPIN_COUNT = 10_000_000;

/**
 * The `spinCount` limit applied by default to untrusted files. Office writes 100,000 and real
 * files never approach 1,000,000; the spin is synchronous, so a hostile file naming 10,000,000
 * rounds would freeze the thread for seconds. Pass `maxSpinCount` (at most {@link MAX_SPIN_COUNT})
 * to accept more.
 */
export const DEFAULT_MAX_SPIN_COUNT = 1_000_000;

/** Limits a caller may relax for a trusted file. */
export interface SpinLimits {
	/** The largest accepted spin count (default {@link DEFAULT_MAX_SPIN_COUNT}, at most {@link MAX_SPIN_COUNT}). */
	maxSpinCount?: number;
}

/** Whether `spinCount` is an integer from 0 up to the effective limit (absent counts as 0). */
export function isValidSpinCount(
	spinCount: number | undefined,
	maxSpinCount: number = DEFAULT_MAX_SPIN_COUNT,
): boolean {
	const limit = Math.min(maxSpinCount, MAX_SPIN_COUNT);
	return (
		spinCount === undefined || (Number.isInteger(spinCount) && spinCount >= 0 && spinCount <= limit)
	);
}

/**
 * The ECMA-376 / [MS-OFFCRYPTO] 2.3.7.1 password spin: H(salt + password as UTF-16LE), then
 * `spinCount` rounds of H(previous + round as uint32 LE). `salt` may be empty (a salt-less
 * verifier simply prepends nothing). Synchronous, so 100,000 rounds take milliseconds.
 */
export function spinPasswordHash(
	password: string,
	salt: Uint8Array,
	spinCount: number,
	digest: DigestFunction,
): Uint8Array {
	const input = new Uint8Array(salt.length + password.length * 2);
	input.set(salt);
	for (let i = 0; i < password.length; i++) {
		const code = password.charCodeAt(i);
		input[salt.length + i * 2] = code & 0xff;
		input[salt.length + i * 2 + 1] = code >> 8;
	}
	let value = digest(input);
	const buffer = new Uint8Array(value.length + 4);
	const view = new DataView(buffer.buffer);
	for (let i = 0; i < spinCount; i++) {
		buffer.set(value);
		view.setUint32(value.length, i, true);
		value = digest(buffer);
	}
	return value;
}

/** The raw hash bytes, or `undefined` when the digest, salt or spin count is unusable. */
function hashBytes(
	password: string,
	{ algorithmName, saltValue, spinCount }: Omit<PasswordHash, 'hashValue'>,
	limits: SpinLimits,
): Uint8Array | undefined {
	const digest = digestFunction(algorithmName);
	const salt = decodeBase64(saltValue);
	if (!digest || !salt || !isValidSpinCount(spinCount, limits.maxSpinCount)) return undefined;
	return spinPasswordHash(password, salt, spinCount ?? 0, digest);
}

/**
 * The base64 ECMA-376 agile hash of `password` for the given salt, spin count and digest, as
 * Excel, Word and PowerPoint store it. `undefined` when the digest is not one {@link digestSync}
 * computes, the salt is not base64, or the spin count is negative, fractional, non-finite or
 * above the limit ({@link DEFAULT_MAX_SPIN_COUNT} unless `limits.maxSpinCount` raises it).
 */
export function hashPassword(
	password: string,
	options: Omit<PasswordHash, 'hashValue'>,
	limits: SpinLimits = {},
): string | undefined {
	const bytes = hashBytes(password, options, limits);
	return bytes && encodeBase64(bytes);
}

/**
 * Whether `password` matches `hash`, comparing decoded bytes (so white space or missing base64
 * padding in the file does not matter). `undefined` when the digest is not one this module
 * computes; `false` (never a throw) for a malformed salt, hash or spin count.
 */
export function verifyPasswordHash(
	password: string,
	hash: PasswordHash,
	limits: SpinLimits = {},
): boolean | undefined {
	if (!digestFunction(hash.algorithmName)) return undefined;
	const expected = decodeBase64(hash.hashValue);
	const actual = expected && hashBytes(password, hash, limits);
	return !!expected && !!actual && timingSafeEqual(actual, expected);
}
