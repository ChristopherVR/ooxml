import type { ModernPasswordHash, SheetProtection, Workbook, Worksheet } from '../model.js';
import { type EditContext, sheetAt } from './context.js';
import { sha512 } from './sha512.js';

/**
 * Excel's legacy 16-bit password verifier (ECMA-376 Part 4, 14.7.1), upper-case hex as the
 * `password` / `workbookPassword` attributes store it. It is a checksum, not encryption.
 */
export function legacyPasswordHash(password: string): string {
	let hash = 0;
	for (let i = 0; i < password.length; i++) {
		const value = (password.charCodeAt(i) & 0xffff) << (i + 1);
		hash ^= (value & 0x7fff) | (value >> 15);
	}
	hash ^= password.length;
	hash ^= 0xce4b;
	return hash.toString(16).toUpperCase();
}

const sameHash = (hash: string | undefined, password: string): boolean =>
	hash === undefined || hash.toUpperCase() === legacyPasswordHash(password);

const protectionOf = (sheet: Worksheet | SheetProtection | undefined) =>
	sheet && 'rows' in sheet ? sheet.protection : sheet;

/**
 * Whether `password` unlocks a sheet's protection (true when it is not protected or has no
 * password). Excel's modern hash is checked when it uses SHA-512 (what Excel writes); a modern
 * hash with another digest needs {@link verifySheetPasswordAsync}, and without a legacy hash to
 * fall back on it never unlocks here, so a sheet is never unprotected by mistake.
 */
export function verifySheetPassword(
	sheet: Worksheet | SheetProtection | undefined,
	password: string,
): boolean {
	const protection = protectionOf(sheet);
	if (!protection?.sheet) return true;
	const modern = protection.modernHash;
	if (!modern) return sameHash(protection.passwordHash, password);
	if (modern.algorithmName.toUpperCase() === 'SHA-512')
		return spinHash(password, modern, sha512) === modern.hashValue;
	return protection.passwordHash !== undefined && sameHash(protection.passwordHash, password);
}

/**
 * Like {@link verifySheetPassword}, but also checks modern hashes that use SHA-1, SHA-256 or
 * SHA-384, through Web Crypto.
 */
export async function verifySheetPasswordAsync(
	sheet: Worksheet | SheetProtection | undefined,
	password: string,
): Promise<boolean> {
	const protection = protectionOf(sheet);
	const modern = protection?.sheet ? protection.modernHash : undefined;
	const digest = modern ? WEB_CRYPTO_DIGESTS[modern.algorithmName.toUpperCase()] : undefined;
	if (!modern || !digest || digest === 'SHA-512') return verifySheetPassword(sheet, password);
	return (await modernPasswordHash(password, modern, digest)) === modern.hashValue;
}

const WEB_CRYPTO_DIGESTS: Readonly<Record<string, string>> = {
	'SHA-1': 'SHA-1',
	'SHA-256': 'SHA-256',
	'SHA-384': 'SHA-384',
	'SHA-512': 'SHA-512',
};

const fromBase64 = (text: string): Uint8Array =>
	Uint8Array.from(atob(text), (c) => c.charCodeAt(0));
const toBase64 = (bytes: Uint8Array): string => btoa(String.fromCharCode(...bytes));

/** The salt followed by the password as UTF-16LE: the input of the first hash round. */
function saltedPassword(password: string, saltValue: string): Uint8Array<ArrayBuffer> {
	const salt = fromBase64(saltValue);
	const input = new Uint8Array(salt.length + password.length * 2);
	input.set(salt);
	for (let i = 0; i < password.length; i++) {
		const code = password.charCodeAt(i);
		input[salt.length + i * 2] = code & 0xff;
		input[salt.length + i * 2 + 1] = code >> 8;
	}
	return input;
}

/** ECMA-376 agile hashing: H(salt + password), then `spinCount` rounds of H(previous + i). */
function spinHash(
	password: string,
	hash: Pick<ModernPasswordHash, 'saltValue' | 'spinCount'>,
	digest: (data: Uint8Array) => Uint8Array,
): string {
	let value = digest(saltedPassword(password, hash.saltValue));
	const buffer = new Uint8Array(value.length + 4);
	const view = new DataView(buffer.buffer);
	for (let i = 0; i < hash.spinCount; i++) {
		buffer.set(value);
		view.setUint32(value.length, i, true);
		value = digest(buffer);
	}
	return toBase64(value);
}

/**
 * The base64 ECMA-376 agile hash of `password` for the salt and iteration count in `hash`.
 * SHA-512 runs synchronously; other digests go through Web Crypto.
 */
export async function modernPasswordHash(
	password: string,
	hash: Pick<ModernPasswordHash, 'saltValue' | 'spinCount'>,
	digest = 'SHA-512',
): Promise<string> {
	if (digest.toUpperCase() === 'SHA-512') return spinHash(password, hash, sha512);
	let value = new Uint8Array(
		await crypto.subtle.digest(digest, saltedPassword(password, hash.saltValue)),
	);
	const buffer = new Uint8Array(value.length + 4);
	const view = new DataView(buffer.buffer);
	for (let i = 0; i < hash.spinCount; i++) {
		buffer.set(value);
		view.setUint32(value.length, i, true);
		value = new Uint8Array(await crypto.subtle.digest(digest, buffer));
	}
	return toBase64(value);
}

/** Whether `password` unlocks the workbook structure protection. */
export function verifyWorkbookPassword(workbook: Workbook, password: string): boolean {
	if (!workbook.structureLocked) return true;
	return sameHash(workbook.workbookPasswordHash, password);
}

const WRONG_PASSWORD = 'The password you supplied is not correct.';

/**
 * Protects (or with `undefined`, unprotects) a sheet. A non-empty `password` sets the legacy
 * hash; `''` removes it. Unprotecting with a password checks it first and throws when wrong.
 */
export function setSheetProtection(
	ctx: EditContext,
	s: number,
	protection: SheetProtection | undefined,
	password?: string,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	if (!protection && password !== undefined && !verifySheetPassword(sheet, password))
		throw new Error(WRONG_PASSWORD);
	ctx.run(
		protection?.sheet ? 'Protect sheet' : 'Unprotect sheet',
		'sheets',
		[{ kind: 'sheet', sheet: s }],
		() => {
			if (!protection) {
				delete sheet.protection;
				return;
			}
			const next = structuredClone(protection);
			// A new password replaces both hashes; the modern one would no longer match it.
			if (password) {
				next.passwordHash = legacyPasswordHash(password);
				delete next.modernHash;
			} else if (password === '') {
				delete next.passwordHash;
				delete next.modernHash;
			}
			sheet.protection = next;
		},
		{ sheet: s },
	);
}

/**
 * Locks or unlocks the workbook structure. Locking with a non-empty password stores its legacy
 * hash; unlocking with a password checks it first and throws when wrong.
 */
export function setWorkbookProtection(ctx: EditContext, locked: boolean, password?: string): void {
	const { workbook } = ctx;
	if (!locked && password !== undefined && !verifyWorkbookPassword(workbook, password))
		throw new Error(WRONG_PASSWORD);
	ctx.run(locked ? 'Protect workbook' : 'Unprotect workbook', 'sheets', [{ kind: 'meta' }], () => {
		if (locked) {
			workbook.structureLocked = true;
			if (password) workbook.workbookPasswordHash = legacyPasswordHash(password);
			else if (password === '') delete workbook.workbookPasswordHash;
		} else {
			delete workbook.structureLocked;
			delete workbook.workbookPasswordHash;
		}
	});
}
