import type { SheetProtection, Workbook, Worksheet } from '../model.js';
import { type EditContext, sheetAt } from './context.js';

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

/**
 * Whether `password` unlocks a sheet's protection (true when it is not protected or has no
 * legacy password). Only the legacy hash is checked: a sheet protected with just Excel's modern
 * SHA-512 hash reports false for every non-empty password.
 */
export function verifySheetPassword(
	sheet: Worksheet | SheetProtection | undefined,
	password: string,
): boolean {
	const protection = sheet && 'rows' in sheet ? sheet.protection : sheet;
	if (!protection?.sheet) return true;
	return sameHash(protection.passwordHash, password);
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
			if (password) next.passwordHash = legacyPasswordHash(password);
			else if (password === '') delete next.passwordHash;
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
