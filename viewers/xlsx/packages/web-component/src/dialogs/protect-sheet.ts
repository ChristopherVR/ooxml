// Protect Sheet and Protect Workbook (structure), with an optional password that the core stores
// as Excel's legacy hash, and the Unprotect prompts that check it. The legacy hash is a checksum,
// not encryption: it keeps honest users from editing by accident, nothing more.
import {
	type SheetProtection,
	verifySheetPassword,
	verifyWorkbookPassword,
} from '@christophervr/xlsx-core';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { checkbox, field, fieldset, invalid, textInput } from './fields.js';
import { showDialog } from './frame.js';

export const PROTECT_ACTIONS: ReadonlyArray<readonly [allow: string, label: string, on: boolean]> =
	[
		['selectLockedCells', 'Select locked cells', true],
		['selectUnlockedCells', 'Select unlocked cells', true],
		['formatCells', 'Format cells', false],
		['formatColumns', 'Format columns', false],
		['formatRows', 'Format rows', false],
		['insertColumns', 'Insert columns', false],
		['insertRows', 'Insert rows', false],
		['insertHyperlinks', 'Insert hyperlinks', false],
		['deleteColumns', 'Delete columns', false],
		['deleteRows', 'Delete rows', false],
		['sort', 'Sort', false],
		['autoFilter', 'Use AutoFilter', false],
		['objects', 'Edit objects', false],
	];

const passwordInput = (ctx: EditorContext): HTMLInputElement => {
	const input = textInput(ctx);
	input.type = 'password';
	input.autocomplete = 'new-password';
	return input;
};

/** A password and its confirmation; `read` is undefined (after a warning) when they differ. */
function passwordPair(ctx: EditorContext, label: string) {
	const first = passwordInput(ctx);
	const again = passwordInput(ctx);
	return {
		fields: [field(ctx, label, first), field(ctx, 'Reenter password to proceed:', again)],
		read: (): string | undefined => {
			if (first.value !== again.value) {
				invalid(ctx, again, 'Confirmation password is not identical.');
				return undefined;
			}
			return first.value;
		},
	};
}

export function openProtectSheet(ctx: EditorContext): Promise<SheetProtection | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const protect = checkbox(ctx, 'Protect worksheet and contents of locked cells', true);
	const password = passwordPair(ctx, 'Password to unprotect sheet:');
	const allowed = t.ws.protection?.allow;
	const checks = PROTECT_ACTIONS.map(
		([allow, label, on]) =>
			[allow, checkbox(ctx, label, allowed ? allowed.includes(allow) : on)] as const,
	);
	return showDialog<SheetProtection>(ctx, {
		name: 'protect-sheet',
		heading: 'Protect Sheet',
		body: [
			protect.wrapper,
			...password.fields,
			fieldset(ctx, 'Allow all users of this worksheet to:', ...checks.map(([, c]) => c.wrapper)),
		],
		opened: () => protect.input.focus(),
		submit: () => {
			if (!protect.input.checked) {
				ctx.toast(ctx.t('Select Protect worksheet and contents of locked cells first.'), 'warning');
				return undefined;
			}
			const secret = password.read();
			if (secret === undefined) return undefined;
			const protection: SheetProtection = {
				sheet: true,
				allow: checks.filter(([, c]) => c.input.checked).map(([allow]) => allow),
			};
			t.session.setSheetProtection(t.sheet, protection, secret);
			return protection;
		},
	});
}

/** Asks for a password until it passes `check`; resolves the password, or undefined. */
function askPassword(
	ctx: EditorContext,
	name: string,
	heading: string,
	check: (password: string) => boolean,
): Promise<string | undefined> {
	const input = passwordInput(ctx);
	return showDialog<string>(ctx, {
		name,
		heading,
		body: [field(ctx, 'Password:', input)],
		opened: () => input.focus(),
		submit: () =>
			check(input.value)
				? input.value
				: invalid(ctx, input, 'The password you supplied is not correct.'),
	});
}

/** Review > Unprotect Sheet: asks for the password when the sheet has one. */
export async function unprotectSheet(ctx: EditorContext): Promise<boolean> {
	const t = target(ctx);
	if (!t?.ws.protection?.sheet) return false;
	if (!t.ws.protection.passwordHash) {
		t.session.setSheetProtection(t.sheet, undefined);
		return true;
	}
	const password = await askPassword(ctx, 'unprotect-sheet', 'Unprotect Sheet', (value) =>
		verifySheetPassword(t.ws, value),
	);
	if (password === undefined) return false;
	t.session.setSheetProtection(t.sheet, undefined, password);
	return true;
}

/** Review > Protect Workbook: locks the structure with an optional password, or unlocks it. */
export async function toggleWorkbookProtection(ctx: EditorContext): Promise<boolean> {
	const session = ctx.session();
	if (!session) return false;
	const workbook = session.workbook;
	if (workbook.structureLocked) {
		if (!workbook.workbookPasswordHash) {
			session.setWorkbookProtection(false);
			return true;
		}
		const password = await askPassword(ctx, 'unprotect-workbook', 'Unprotect Workbook', (value) =>
			verifyWorkbookPassword(workbook, value),
		);
		if (password === undefined) return false;
		session.setWorkbookProtection(false, password);
		return true;
	}
	const password = passwordPair(ctx, 'Password (optional):');
	const structure = checkbox(ctx, 'Structure', true);
	structure.input.disabled = true;
	const locked = await showDialog<true>(ctx, {
		name: 'protect-workbook',
		heading: 'Protect Structure and Windows',
		body: [fieldset(ctx, 'Protect workbook for', structure.wrapper), ...password.fields],
		submit: () => {
			const secret = password.read();
			if (secret === undefined) return undefined;
			session.setWorkbookProtection(true, secret);
			return true;
		},
	});
	return locked === true;
}
