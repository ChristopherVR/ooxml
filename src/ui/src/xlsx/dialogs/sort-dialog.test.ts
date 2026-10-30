// @vitest-environment jsdom
import { createWorkbook, getCell } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	pressKey,
	setValue,
} from '../commands/test-support';
import { clearUserSortLists, userSortLists } from './custom-lists';
import { registerRuleDialogs } from './register-rules';

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

	it('sorts on a custom list with Options for case and orientation', async () => {
		const ctx = createTestContext(createWorkbook());
		registerRuleDialogs(ctx);
		ctx.session()!.setRangeValues(0, { row: 0, col: 0 }, [['Day'], ['Wed'], ['Mon'], ['Sun']]);
		ctx.select('A1');
		const result = ctx.dialogs.open('sort');
		const d = dialogEl(ctx, 'sort');
		setValue(selects(d, 'Order')[0]!, 'custom');
		const pick = selects(d, 'Custom List...')[0]!;
		expect(pick.hidden).toBe(false);
		setValue(pick, '0');
		clickButton(d, 'OK');
		const keys = (await result) as readonly { customList?: readonly string[] }[] | undefined;
		expect(keys?.[0]?.customList).toEqual(['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']);
		expect(column(ctx, 0)).toEqual(['Day', 'Sun', 'Mon', 'Wed']);
	});

	it('offers colours and orders by them, and switches to rows left to right', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('sort');
		const d = dialogEl(ctx, 'sort');
		setValue(selects(d, 'Sort On')[0]!, 'cellColor');
		expect([...selects(d, 'Order')[0]!.options].map((o) => o.value)).toEqual(['top', 'bottom']);
		clickButton(d, 'Options...');
		setValue(selects(d, 'Orientation')[0]!, 'lr');
		expect(selects(d, 'Row')).toHaveLength(1);
		expect([...selects(d, 'Order')[0]!.options].map((o) => o.textContent)).toEqual([
			'On Left',
			'On Right',
		]);
		setValue(selects(d, 'Sort On')[0]!, 'value');
		setValue(selects(d, 'Row')[0]!, '1');
		clickButton(d, 'OK');
		await result;
		expect(ctx.workbook()!.sheets[0]!.sortState?.columnSort).toBe(true);
	});

	it('keeps a typed custom list for later dialogs', async () => {
		clearUserSortLists();
		const ctx = setup();
		const result = ctx.dialogs.open('sort');
		const d = dialogEl(ctx, 'sort');
		setValue(selects(d, 'Order')[0]!, 'custom');
		setValue(selects(d, 'Custom List...')[0]!, 'new');
		setValue(d.querySelector<HTMLInputElement>('input[aria-label="List entries"]')!, 'b, a, c');
		clickButton(d, 'OK');
		expect(
			((await result) as readonly { customList?: readonly string[] }[] | undefined)?.[0]
				?.customList,
		).toEqual(['b', 'a', 'c']);
		expect(userSortLists()).toEqual([['b', 'a', 'c']]);
		void ctx.dialogs.open('sort');
		const again = [...selects(dialogEl(ctx, 'sort'), 'Custom List...')[0]!.options];
		expect(again.map((o) => o.textContent)).toContain('b, a, c');
		clearUserSortLists();
	});
});
