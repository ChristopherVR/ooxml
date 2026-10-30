// The password Word hashes into `w:documentProtection` in its own spelling (`w:cryptAlgorithmSid`,
// `w:hash`, `w:salt`): the password is first reduced to the 4-byte key of the old binary Word
// formats (ECMA-376 Part 4, 14.8.1; [MS-OFFCRYPTO] 2.3.7.1 and 2.3.7.4), written as 8 upper-case hex
// characters, and that text is what the salted spin hashes. Checked against documents saved by
// desktop Word 16 (src/core/docx/__fixtures__/protection).
import {
	decodeBase64,
	digestFunction,
	isValidSpinCount,
	spinPasswordHash,
	timingSafeEqual,
} from '../digest/index';
import type { ProtectionPassword } from './form-field-model';

/** The initial code array of the key derivation, indexed by password length - 1. */
const INITIAL_CODE_ARRAY = [
	0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9,
	0x280c, 0xa96a, 0x4ec3,
];

/** The encryption matrix: one row per password position, counted from the end of 15 characters. */
const ENCRYPTION_MATRIX = [
	[0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09],
	[0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf],
	[0x4563, 0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68, 0x5ad0],
	[0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40],
	[0xd849, 0xa0b3, 0x5147, 0xa28e, 0x553d, 0xaa7a, 0x44d5],
	[0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496, 0x390d, 0x721a],
	[0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe, 0xa7fc, 0x5fd9],
	[0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0],
	[0xb861, 0x60e3, 0xc1c6, 0x93ad, 0x377b, 0x6ef6, 0xddec],
	[0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10],
	[0xaa51, 0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168],
	[0x76b4, 0xed68, 0xcaf1, 0x85c3, 0x1ba7, 0x374e, 0x6e9c],
	[0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6, 0x1dad],
	[0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373, 0x06e6, 0x0dcc],
	[0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48c4],
];

/** The 15-bit rotate left of the verifier derivation. */
const rotate = (value: number): number =>
	((value & 0x4000) !== 0 ? 1 : 0) | ((value << 1) & 0x7fff);

/**
 * The key the old binary formats derive from a password, as the 8 upper-case hex characters Word
 * hashes: the verifier word then the key word, each low byte first. The password is cut to 15
 * characters, and each character gives its low byte, or its high byte when that is zero.
 */
export function legacyPasswordKeyText(password: string): string {
	const bytes = Array.from(password.slice(0, 15), (char) => {
		const code = char.charCodeAt(0);
		return code & 0xff || (code >> 8) & 0xff;
	});
	if (!bytes.length) return '00000000';
	let high = INITIAL_CODE_ARRAY[bytes.length - 1]!;
	let row = 15 - bytes.length;
	for (let byte of bytes) {
		for (const word of ENCRYPTION_MATRIX[row++]!) {
			if (byte & 1) high ^= word;
			byte >>= 1;
		}
	}
	let verifier = 0;
	for (let index = bytes.length - 1; index >= 0; index--)
		verifier = rotate(verifier) ^ bytes[index]!;
	verifier = (rotate(verifier) ^ bytes.length ^ 0xce4b) & 0xffff;
	const hex = (value: number) => value.toString(16).padStart(2, '0');
	return (
		hex(verifier & 0xff) +
		hex(verifier >> 8) +
		hex(high & 0xff) +
		hex((high >> 8) & 0xff)
	).toUpperCase();
}

/**
 * Whether `password` matches a verifier in Word's spelling; `undefined` when its digest, salt or
 * spin count cannot be used.
 */
export function verifyLegacyProtectionPassword(
	password: string,
	verifier: ProtectionPassword,
): boolean | undefined {
	const digest = digestFunction(verifier.algorithmName);
	const salt = decodeBase64(verifier.saltValue);
	const expected = decodeBase64(verifier.hashValue);
	if (!digest || !salt || !isValidSpinCount(verifier.spinCount)) return undefined;
	if (!expected) return false;
	const actual = spinPasswordHash(
		legacyPasswordKeyText(password),
		salt,
		verifier.spinCount ?? 0,
		digest,
	);
	return timingSafeEqual(actual, expected);
}
