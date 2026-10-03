import { decodeBase64, encodeBase64 } from './base64.js';
import { digestFunction } from './digest.js';

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
 * The largest `spinCount` honoured. Excel writes 100,000; anything far beyond that is a
 * malformed or hostile file, and hashing it would freeze the thread.
 */
export const MAX_SPIN_COUNT = 10_000_000;

const validSpinCount = (spinCount: number | undefined): spinCount is number | undefined =>
	spinCount === undefined ||
	(Number.isInteger(spinCount) && spinCount >= 0 && spinCount <= MAX_SPIN_COUNT);

/** H(salt + password as UTF-16LE), then `spinCount` rounds of H(previous + round as uint32 LE). */
function spin(
	password: string,
	salt: Uint8Array,
	spinCount: number,
	digest: (data: Uint8Array) => Uint8Array,
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
): Uint8Array | undefined {
	const digest = digestFunction(algorithmName);
	const salt = decodeBase64(saltValue);
	if (!digest || !salt || !validSpinCount(spinCount)) return undefined;
	return spin(password, salt, spinCount ?? 0, digest);
}

/**
 * The base64 ECMA-376 agile hash of `password` for the given salt, spin count and digest, as
 * Excel, Word and PowerPoint store it. `undefined` when the digest is not one {@link digestSync}
 * computes, the salt is not base64, or the spin count is negative, fractional, non-finite or
 * above {@link MAX_SPIN_COUNT}.
 */
export function hashPassword(
	password: string,
	options: Omit<PasswordHash, 'hashValue'>,
): string | undefined {
	const bytes = hashBytes(password, options);
	return bytes && encodeBase64(bytes);
}

/**
 * Whether `password` matches `hash`, comparing decoded bytes (so white space or missing base64
 * padding in the file does not matter). `undefined` when the digest is not one this module
 * computes; `false` (never a throw) for a malformed salt, hash or spin count.
 */
export function verifyPasswordHash(password: string, hash: PasswordHash): boolean | undefined {
	if (!digestFunction(hash.algorithmName)) return undefined;
	const expected = decodeBase64(hash.hashValue);
	const actual = expected && hashBytes(password, hash);
	if (!expected || !actual || actual.length !== expected.length) return false;
	let difference = 0;
	for (let i = 0; i < actual.length; i++) difference |= actual[i]! ^ expected[i]!;
	return difference === 0;
}
