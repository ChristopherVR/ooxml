// @vitest-environment jsdom
import { createWorkbook, getCell } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
} from '../commands/test-support.js';
import { searchFunctions } from './insert-function.js';
import { registerNavigationDialogs } from './register-navigation.js';

afterEach(() => (document.body.innerHTML = ''));

describe('insert function dialog', () => {
	it('searches by name and description', () => {
		expect(searchFunctions('vlookup')[0]?.name).toBe('VLOOKUP');
		expect(searchFunctions('').length).toBeGreaterThan(100);
	});

	it('filters by category, shows syntax, and writes =NAME(args) without a grid', async () => {
		const ctx = createTestContext(createWorkbook());
		registerNavigationDialogs(ctx);
		ctx.select('C3');
		const done = ctx.dialogs.open('insert-function');
		const dialog = dialogEl(ctx, 'insert-function');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Or select a category:'), 'Math & Trig');
		const option = [...dialog.querySelectorAll<HTMLElement>('[role="option"]')].find(
			(o) => o.dataset.value === 'SUM',
		)!;
		option.click();
		expect(dialog.querySelector('.xve-syntax')?.textContent).toContain('SUM(');
		setValue(inputByLabel(dialog, 'Arguments (optional):'), '1,2');
		clickButton(dialog, 'OK');
		expect(await done).toBe('SUM');
		const cell = getCell(ctx.workbook()!.sheets[0]!, 2, 2);
		expect(cell?.formula).toBe('SUM(1,2)');
		expect(cell?.value).toBe(3);
	});

	it('starts the cell editor with =NAME( when a grid is mounted', async () => {
		const beginEdit = vi.fn();
		const ctx = createTestContext(createWorkbook(), { grid: { beginEdit } });
		registerNavigationDialogs(ctx);
		const done = ctx.dialogs.open('insert-function', { fn: 'IF' });
		const list = dialogEl(ctx, 'insert-function').querySelector<HTMLElement>('[role="listbox"]')!;
		pressKey(list, 'Enter');
		expect(await done).toBe('IF');
		expect(beginEdit).toHaveBeenCalledWith('=IF(');
	});

	it('Go runs a search; Cancel and Escape insert nothing', async () => {
		const ctx = createTestContext(createWorkbook());
		registerNavigationDialogs(ctx);
		const first = ctx.dialogs.open('insert-function');
		const dialog = dialogEl(ctx, 'insert-function');
		setValue(inputByLabel(dialog, 'Search for a function:'), 'xlookup');
		clickButton(dialog, 'Go');
		expect(dialog.querySelector('[role="option"]')?.textContent).toBe('XLOOKUP');
		clickButton(dialog, 'Cancel');
		expect(await first).toBeUndefined();
		const second = ctx.dialogs.open('insert-function');
		pressKey(inputByLabel(dialogEl(ctx, 'insert-function'), 'Arguments (optional):'), 'Escape');
		expect(await second).toBeUndefined();
		expect(ctx.workbook()!.sheets[0]!.rows.size).toBe(0);
	});
});
