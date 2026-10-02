// @vitest-environment jsdom
import { createWorkbook, getCell } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	pressKey,
	setValue,
} from '../commands/test-support.js';
import { registerRuleDialogs } from './register-rules.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook());
	registerRuleDialogs(ctx);
	ctx.session()!.setRangeValues(0, { row: 0, col: 0 }, [
		['Name', 'Score'],
		['b', 2],
		['a', 2],
		['c', 1],
	]);
	ctx.select('A1');
	return ctx;
}
const column = (ctx: ReturnType<typeof setup>, col: number) =>
	[0, 1, 2, 3].map((r) => getCell(ctx.workbook()!.sheets[0]!, r, col)?.value);
const selects = (d: HTMLElement, label: string) => [
	...d.querySelectorAll<HTMLSelectElement>(`select[aria-label="${label}"]`),
];

describe('sort dialog', () => {
	it('detects headers and lists header names', () => {
		const ctx = setup();
		void ctx.dialogs.open('sort');
		const d = dialogEl(ctx, 'sort');
		expect(
			d.querySelector<HTMLInputElement>('input[aria-label="My data has headers"]')!.checked,
		).toBe(true);
		expect([...selects(d, 'Column')[0]!.options].map((o) => o.textContent)).toEqual([
			'Name',
			'Score',
		]);
	});

	it('sorts by several levels', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('sort');
		const d = dialogEl(ctx, 'sort');
		setValue(selects(d, 'Column')[0]!, '1');
		setValue(selects(d, 'Order')[0]!, 'desc');
		clickButton(d, 'Add Level');
		setValue(selects(d, 'Column')[1]!, '0');
		clickButton(d, 'OK');
		expect(await result).toEqual([
			{ col: 1, descending: true },
			{ col: 0, descending: false },
		]);
		expect(column(ctx, 0)).toEqual(['Name', 'a', 'b', 'c']);
	});

	it('moves and deletes levels; Cancel and Escape do not sort', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('sort');
		const d = dialogEl(ctx, 'sort');
		clickButton(d, 'Copy Level');
		expect(selects(d, 'Column')).toHaveLength(2);
		clickButton(d, 'Move Up');
		clickButton(d, 'Delete Level');
		expect(selects(d, 'Column')).toHaveLength(1);
		clickButton(d, 'Cancel');
		expect(await result).toBeUndefined();
		const again = ctx.dialogs.open('sort');
		pressKey(selects(dialogEl(ctx, 'sort'), 'Column')[0]!, 'Escape');
		expect(await again).toBeUndefined();
		expect(column(ctx, 0)).toEqual(['Name', 'b', 'a', 'c']);
	});
});
