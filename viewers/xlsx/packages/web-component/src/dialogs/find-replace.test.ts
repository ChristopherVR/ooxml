// @vitest-environment jsdom
import { createWorkbook, getCell } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
} from '../commands/test-support.js';
import { orderFrom } from './find-replace.js';
import { registerNavigationDialogs } from './register-navigation.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook({ sheets: ['Sheet1', 'Sheet2'] }));
	registerNavigationDialogs(ctx);
	const s = ctx.session()!;
	s.setCellInput(0, 0, 0, 'apple');
	s.setCellInput(0, 2, 1, 'Apple pie');
	s.setCellInput(0, 4, 0, 'pear');
	s.setCellInput(1, 1, 1, 'apple');
	return ctx;
}

describe('find and replace dialog', () => {
	it('orders matches after the active cell, wrapping', () => {
		const m = (sheet: number, row: number, col: number) => ({ sheet, row, col, text: '' });
		expect(
			orderFrom([m(0, 0, 0), m(0, 2, 1), m(1, 0, 0)], 0, 1, 0, false).map((x) => x.row),
		).toEqual([2, 0, 0]);
	});

	it('Find Next selects the next match on the sheet', async () => {
		const ctx = setup();
		const done = ctx.dialogs.open('find-replace');
		const dialog = dialogEl(ctx, 'find-replace');
		setValue(inputByLabel(dialog, 'Find what:'), 'apple');
		clickButton(dialog, 'Find Next');
		expect(ctx.selection.get().active).toEqual({ row: 2, col: 1 });
		clickButton(dialog, 'Find Next');
		expect(ctx.selection.get().active).toEqual({ row: 0, col: 0 });
		clickButton(dialog, 'Close');
		expect(await done).toBeUndefined();
	});

	it('Find All lists matches across the workbook and selecting one switches sheets', () => {
		const ctx = setup();
		void ctx.dialogs.open('find-replace');
		const dialog = dialogEl(ctx, 'find-replace');
		setValue(inputByLabel(dialog, 'Find what:'), 'apple');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Within:'), 'workbook');
		inputByLabel(dialog, 'Match entire cell contents').click();
		clickButton(dialog, 'Find All');
		const rows = dialog.querySelectorAll('.xve-results tbody tr');
		expect(rows.length).toBe(2);
		(rows[1] as HTMLElement).click();
		expect(ctx.activeSheet()).toBe(1);
		expect(ctx.selection.get().active).toEqual({ row: 1, col: 1 });
	});

	it('Replace All changes the matching cells and reports the count', async () => {
		const ctx = setup();
		const done = ctx.dialogs.open('find-replace', { tab: 'replace' });
		const dialog = dialogEl(ctx, 'find-replace');
		setValue(inputByLabel(dialog, 'Find what:'), 'apple');
		setValue(inputByLabel(dialog, 'Replace with:'), 'plum');
		clickButton(dialog, 'Replace All');
		const ws = ctx.workbook()!.sheets[0]!;
		expect(getCell(ws, 0, 0)?.value).toBe('plum');
		expect(getCell(ws, 2, 1)?.value).toBe('plum pie');
		expect(getCell(ctx.workbook()!.sheets[1]!, 1, 1)?.value).toBe('apple');
		expect(ctx.toasts.at(-1)?.message).toBe('All done. We made 2 replacements.');
		pressKey(inputByLabel(dialog, 'Find what:'), 'Escape');
		expect(await done).toBe(2);
	});

	it('Replace replaces the current match and moves on', () => {
		const ctx = setup();
		void ctx.dialogs.open('find-replace', { tab: 'replace' });
		const dialog = dialogEl(ctx, 'find-replace');
		setValue(inputByLabel(dialog, 'Find what:'), 'apple');
		setValue(inputByLabel(dialog, 'Replace with:'), 'fig');
		clickButton(dialog, 'Find Next');
		clickButton(dialog.querySelector<HTMLElement>('.xve-dialog-footer')!, 'Replace');
		const ws = ctx.workbook()!.sheets[0]!;
		expect(getCell(ws, 2, 1)?.value).toBe('fig pie');
		expect(getCell(ws, 0, 0)?.value).toBe('apple');
	});

	it('opens on Find and disables Replace in read-only mode', () => {
		const ctx = setup();
		ctx.setReadOnly(true);
		void ctx.dialogs.open('find-replace', { tab: 'replace' });
		const tab = dialogEl(ctx, 'find-replace').querySelector<HTMLButtonElement>(
			'[data-tab="replace"]',
		)!;
		expect(tab.disabled).toBe(true);
		expect(tab.getAttribute('aria-selected')).toBe('false');
	});

	it('Escape closes without changes', async () => {
		const ctx = setup();
		const done = ctx.dialogs.open('find-replace');
		pressKey(inputByLabel(dialogEl(ctx, 'find-replace'), 'Find what:'), 'Escape');
		expect(await done).toBeUndefined();
		expect(ctx.root.querySelector('[data-dialog]')).toBeNull();
	});
});
