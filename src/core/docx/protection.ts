// Document protection (`w:documentProtection`): reading its password verifier, checking a typed
// password against it and making a new one. Word writes the verifier as `w:cryptAlgorithmSid`,
// `w:hash`, `w:salt` and `w:cryptSpinCount`; ISO 29500 also allows `w:algorithmName`,
// `w:hashValue`, `w:saltValue` and `w:spinCount`. Both use the ECMA-376 password spin from
// `ooxml-core/digest`, the one the Excel sheet protection uses.
import { hashPassword, verifyPasswordHash } from '../digest/index';
import type { DocumentProtection, ProtectionPassword } from './form-field-model';
import { verifyLegacyProtectionPassword } from './protection-legacy';
import { getW, type XmlElement } from './xml';

/** `w:cryptAlgorithmSid` values and the digest each names (ECMA-376 Part 4, 17.15.1.29). */
const DIGEST_BY_SID: Readonly<Record<number, string>> = {
	1: 'MD2',
	2: 'MD4',
	3: 'MD5',
	4: 'SHA-1',
	7: 'RIPEMD-128',
	8: 'RIPEMD-160',
	12: 'SHA-256',
	13: 'SHA-384',
	14: 'SHA-512',
};

const SPIN_COUNT = 100_000;

const integer = (value: string | undefined): number | undefined => {
	const number = value === undefined ? NaN : Number(value);
	return Number.isInteger(number) ? number : undefined;
};

/**
 * The password verifier of a `w:documentProtection` element. `unverifiable` is set when the
 * element has a hash this code cannot check (Word 97's 4-byte hash, an unknown algorithm).
 */
export function readProtectionPassword(element: XmlElement): {
	password?: ProtectionPassword;
	unverifiable?: true;
} {
	const hashValue = getW(element, 'hashValue');
	const saltValue = getW(element, 'saltValue');
	const algorithmName = getW(element, 'algorithmName');
	if (hashValue && saltValue && algorithmName) {
		const spinCount = integer(getW(element, 'spinCount'));
		return {
			password: {
				algorithmName,
				hashValue,
				saltValue,
				...(spinCount !== undefined && { spinCount }),
			},
		};
	}
	const hash = getW(element, 'hash');
	if (!hash) return {};
	const sid = integer(getW(element, 'cryptAlgorithmSid'));
	const salt = getW(element, 'salt');
	const digest = sid === undefined ? undefined : DIGEST_BY_SID[sid];
	if (sid === undefined || !digest || !salt) return { unverifiable: true };
	const spinCount = integer(getW(element, 'cryptSpinCount'));
	return {
		password: {
			algorithmName: digest,
			hashValue: hash,
			saltValue: salt,
			...(spinCount !== undefined && { spinCount }),
			word: { algorithmSid: sid },
		},
	};
}

/**
 * Whether `password` stops the protection: true when there is no password to check, false for a
 * wrong password and `undefined` when the verifier uses a digest or form this code cannot check
 * (the protection then cannot be stopped, so a document is never unprotected by mistake).
 */
export function verifyProtectionPassword(
	protection: DocumentProtection | undefined,
	password: string,
): boolean | undefined {
	if (protection?.unverifiablePassword) return undefined;
	if (!protection?.password) return true;
	// Word's own `cryptAlgorithmSid` hash spins the old binary formats' key of the password.
	if (protection.password.word)
		return verifyLegacyProtectionPassword(password, protection.password);
	return verifyPasswordHash(password, protection.password);
}

/** A new random salt, base64. */
function randomSalt(): string {
	const bytes = new Uint8Array(16);
	globalThis.crypto.getRandomValues(bytes);
	let binary = '';
	for (const byte of bytes) binary += String.fromCharCode(byte);
	return btoa(binary);
}

/**
 * A verifier for `password` in the ISO 29500 form (`w:algorithmName`, `w:hashValue`,
 * `w:saltValue`, `w:spinCount`): SHA-512, 100,000 rounds, a random salt. Word unprotects a
 * document with it (checked with desktop Word through COM), so it can be used to stop protection
 * in Word. `undefined` for an empty password, which Word treats as none.
 */
export function createProtectionPassword(
	password: string,
	saltValue: string = randomSalt(),
): ProtectionPassword | undefined {
	if (!password) return undefined;
	const options = { algorithmName: 'SHA-512', saltValue, spinCount: SPIN_COUNT };
	const hashValue = hashPassword(password, options);
	return hashValue ? { ...options, hashValue } : undefined;
}

/** Starts enforcing `edit` (with `password`, if any), keeping the formatting restriction. */
export function enforceProtection(
	edit: DocumentProtection['edit'],
	options: { password?: string; formatting?: boolean; salt?: string } = {},
): DocumentProtection {
	const password = options.password
		? createProtectionPassword(options.password, options.salt)
		: undefined;
	return {
		edit,
		enforced: true,
		...(options.formatting && { formatting: true }),
		...(password && { password }),
	};
}
