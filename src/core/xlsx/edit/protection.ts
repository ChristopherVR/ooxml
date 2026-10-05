import type { ModernPasswordHash, SheetProtection, Workbook, Worksheet } from '../model.js';
import { type EditContext, sheetAt } from './context.js';
import { hashPassword, verifyPasswordHash } from '../../digest/index.js';

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

const sameLegacyHash = (hash: string, password: string): boolean =>
	hash.toUpperCase() === legacyPasswordHash(password);

/**
 * Whether `password` matches a protection's hashes. The modern hash wins when its digest is one
 * `ooxml-core/digest` computes (SHA-1, SHA-256, SHA-384, SHA-512); with another digest the legacy
 * hash decides, and without one nothing unlocks. No hash at all means no password.
 */
function matchesHashes(
	modern: ModernPasswordHash | undefined,
	legacy: string | undefined,
	password: string,
): boolean {
	const verdict = modern ? verifyPasswordHash(password, modern) : undefined;
	if (verdict !== undefined) return verdict;
	if (legacy !== undefined) return sameLegacyHash(legacy, password);
	return !modern;
}

const protectionOf = (sheet: Worksheet | SheetProtection | undefined) =>
	sheet && 'rows' in sheet ? sheet.protection : sheet;

/**
 * Whether `password` unlocks a sheet's protection (true when it is not protected or has no
 * password). Excel's modern hash is checked synchronously for SHA-1, SHA-256, SHA-384 and SHA-512;
 * a modern hash with any other digest falls back to the legacy hash and, without one, never
 * unlocks, so a sheet is never unprotected by mistake. Malformed hashes fail rather than throw.
 */
export function verifySheetPassword(
	sheet: Worksheet | SheetProtection | undefined,
	password: string,
): boolean {
	const protection = protectionOf(sheet);
	if (!protection?.sheet) return true;
	return matchesHashes(protection.modernHash, protection.passwordHash, password);
}

/** {@link verifySheetPassword} as a promise, kept for callers written against the async API. */
export async function verifySheetPasswordAsync(
	sheet: Worksheet | SheetProtection | undefined,
	password: string,
): Promise<boolean> {
	return verifySheetPassword(sheet, password);
}

/**
 * The base64 ECMA-376 agile hash of `password` for the salt and iteration count in `hash`, as a
 * promise for callers written against the async API; `hashPassword` in `ooxml-core/digest` is the
 * synchronous form. Rejects for a digest other than SHA-1/256/384/512 or unusable parameters.
 */
export async function modernPasswordHash(
	password: string,
	hash: Pick<ModernPasswordHash, 'saltValue' | 'spinCount'>,
	digest = 'SHA-512',
): Promise<string> {
	const value = hashPassword(password, { ...hash, algorithmName: digest });
	if (value === undefined) throw new Error(`Cannot compute a ${digest} password hash`);
	return value;
}

/** Whether `password` unlocks the workbook structure protection (modern or legacy hash). */
export function verifyWorkbookPassword(workbook: Workbook, password: string): boolean {
	if (!workbook.structureLocked) return true;
	return matchesHashes(workbook.workbookModernHash, workbook.workbookPasswordHash, password);
}

const WRONG_PASSWORD = 'The password you supplied is not correct.';

/**
 * Protects (or with `undefined`, unprotects) a sheet. A non-empty `password` sets the legacy
 * hash and drops the modern one (it would no longer match); `''` removes both hashes, and
 * `undefined` keeps whatever `protection` carries. Unprotecting with a password checks it first
 * (against the modern or legacy hash) and throws when it is wrong.
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
 * hash and drops the modern one; `''` removes both, `undefined` keeps them. Unlocking with a
 * password checks it first (against the modern or legacy hash) and throws when it is wrong.
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
			if (password !== undefined) delete workbook.workbookModernHash;
		} else {
			delete workbook.structureLocked;
			delete workbook.workbookPasswordHash;
			delete workbook.workbookModernHash;
		}
	});
}
