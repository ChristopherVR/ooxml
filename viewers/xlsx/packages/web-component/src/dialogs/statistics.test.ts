// @vitest-environment jsdom
import { createWorkbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import { clickButton, createTestContext, dialogEl, pressKey } from '../commands/test-support.js';
import { registerNavigationDialogs } from './register-navigation.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook({ sheets: ['Sheet1', 'Sheet2'] }));
	registerNavigationDialogs(ctx);
	const s = ctx.session()!;
	s.setCellInput(0, 0, 0, '1');
	s.setCellInput(0, 3, 2, '=A1*2');
	s.setCellInput(1, 0, 0, 'x');
	s.setComment(0, { row: 3, col: 2 }, 'Second', 'Ann');
	s.setComment(0, { row: 0, col: 0 }, 'First', 'Bob');
	return ctx;
}

const cells = (el: Element) =>
	[...el.querySelectorAll('tr')].map((r) => [...r.children].map((c) => c.textContent));

describe('workbook statistics', () => {
	it('counts the current sheet and the workbook', async () => {
		const ctx = setup();
		const done = ctx.dialogs.open('workbook-statistics');
		const [sheet, book] = [...dialogEl(ctx, 'workbook-statistics').querySelectorAll('table')];
		expect(cells(sheet!)).toContainEqual(['End of sheet', 'C4']);
		expect(cells(sheet!)).toContainEqual(['Formulas', '1']);
		expect(cells(sheet!)).toContainEqual(['Notes', '2']);
		expect(cells(book!)).toContainEqual(['Sheets', '2']);
		expect(cells(book!)).toContainEqual(['Cells with data', '3']);
		expect(dialogEl(ctx, 'workbook-statistics').querySelector('.xve-btn-primary')).toBeNull();
		clickButton(dialogEl(ctx, 'workbook-statistics'), 'Close');
		expect(await done).toBeUndefined();
	});
});

describe('comments list', () => {
	it('lists comments in order and selects the activated one', async () => {
		const ctx = setup();
		const done = ctx.dialogs.open('comments-list');
		const options = [
			...dialogEl(ctx, 'comments-list').querySelectorAll<HTMLElement>('[role="option"]'),
		];
		expect(options.map((o) => o.textContent)).toEqual(['A1  Bob: First', 'C4  Ann: Second']);
		const list = dialogEl(ctx, 'comments-list').querySelector<HTMLElement>('[role="listbox"]')!;
		pressKey(list, 'ArrowDown');
		pressKey(list, 'Enter');
		expect(await done).toBe('C4');
		expect(ctx.selection.get().active).toEqual({ row: 3, col: 2 });
	});

	it('Escape closes without moving', async () => {
		const ctx = setup();
		const done = ctx.dialogs.open('comments-list');
		pressKey(
			dialogEl(ctx, 'comments-list').querySelector<HTMLElement>('[role="listbox"]')!,
			'Escape',
		);
		expect(await done).toBeUndefined();
		expect(ctx.selection.get().active).toEqual({ row: 0, col: 0 });
	});
});

describe('help dialogs', () => {
	it('shows shortcut rows from the shell and the feature status', async () => {
		const ctx = setup();
		ctx.commands.register({
			id: 'home.bold',
			label: 'Bold',
			shortcut: 'Ctrl+B',
			run: () => undefined,
		});
		const help = ctx.dialogs.open('shortcut-help');
		const rows = cells(dialogEl(ctx, 'shortcut-help').querySelector('tbody')!);
		expect(rows.some((r) => r[0] === 'Bold')).toBe(true);
		clickButton(dialogEl(ctx, 'shortcut-help'), 'Close');
		await help;
		const about = ctx.dialogs.open('feature-status');
		const text = dialogEl(ctx, 'feature-status').textContent ?? '';
		expect(text).toContain('Pivot tables and sparklines');
		expect(text).toContain('Freeze panes');
		pressKey(dialogEl(ctx, 'feature-status').querySelector('button')!, 'Escape');
		expect(await about).toBeUndefined();
	});
});
