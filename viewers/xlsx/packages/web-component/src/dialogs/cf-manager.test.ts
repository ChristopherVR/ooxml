// @vitest-environment jsdom
import { type ConditionalRule, createWorkbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
	tick,
} from '../commands/test-support.js';
import { registerRuleDialogs } from './register-rules.js';

afterEach(() => (document.body.innerHTML = ''));

const A = { start: { row: 0, col: 0 }, end: { row: 9, col: 0 } };
const rule = (formula: string): ConditionalRule => ({
	type: 'expression',
	formula,
	style: {},
	priority: 1,
});

function setup() {
	const ctx = createTestContext(createWorkbook());
	registerRuleDialogs(ctx);
	const session = ctx.session()!;
	session.addConditionalFormat(0, { ranges: [A], rules: [rule('first')] });
	session.addConditionalFormat(0, { ranges: [A], rules: [rule('second')] });
	session.addConditionalFormat(0, { ranges: [A], rules: [rule('third')] });
	ctx.select('A1');
	return ctx;
}
const order = (ctx: ReturnType<typeof setup>) =>
	ctx
		.workbook()!
		.sheets[0]!.conditionalFormats.flatMap((f) => f.rules)
		.sort((a, b) => a.priority - b.priority)
		.map((r) => (r.type === 'expression' ? r.formula : r.type));
const rowOf = (dialog: HTMLElement, text: string) =>
	[...dialog.querySelectorAll<HTMLElement>('tr[data-index]')].find((tr) =>
		tr.textContent?.includes(text),
	)!;

describe('rules manager', () => {
	it('lists rules in priority order', () => {
		const ctx = setup();
		void ctx.dialogs.open('cf-manager');
		const rows = dialogEl(ctx, 'cf-manager').querySelectorAll('tr[data-index]');
		expect([...rows].map((r) => r.textContent?.includes('=third'))).toEqual([true, false, false]);
	});

	it('reorders and deletes as one undo step on OK', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('cf-manager');
		const dialog = dialogEl(ctx, 'cf-manager');
		rowOf(dialog, '=first').click();
		clickButton(dialogEl(ctx, 'cf-manager'), 'Move Up');
		rowOf(dialogEl(ctx, 'cf-manager'), '=third').click();
		clickButton(dialogEl(ctx, 'cf-manager'), 'Delete Rule');
		clickButton(dialogEl(ctx, 'cf-manager'), 'OK');
		await result;
		expect(order(ctx)).toEqual(['first', 'second']);
		ctx.session()!.undo();
		expect(order(ctx)).toEqual(['third', 'second', 'first']);
	});

	it('edits Applies to and Stop If True, and duplicates', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('cf-manager');
		const dialog = dialogEl(ctx, 'cf-manager');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Show formatting rules for:'), 'sheet');
		const row = rowOf(dialogEl(ctx, 'cf-manager'), '=second');
		setValue(row.querySelector<HTMLInputElement>('input.xve-input')!, '$C$1:$C$4 E5');
		row.querySelector<HTMLInputElement>('input[type="checkbox"]')!.click();
		row.click();
		clickButton(dialogEl(ctx, 'cf-manager'), 'Duplicate Rule');
		clickButton(dialogEl(ctx, 'cf-manager'), 'OK');
		await result;
		const formats = ctx.workbook()!.sheets[0]!.conditionalFormats;
		expect(formats).toHaveLength(4);
		const second = formats.filter(
			(f) => f.rules[0]?.type === 'expression' && f.rules[0].formula === 'second',
		);
		expect(second).toHaveLength(2);
		expect(second[0]!.ranges).toEqual([
			{ start: { row: 0, col: 2 }, end: { row: 3, col: 2 } },
			{ start: { row: 4, col: 4 }, end: { row: 4, col: 4 } },
		]);
		expect(second[0]!.rules[0]).toMatchObject({ stopIfTrue: true });
		expect(formats.map((f) => f.rules[0]!.priority)).toEqual([1, 2, 3, 4]);
	});

	it('rejects an invalid Applies to; Cancel and Escape change nothing', async () => {
		const ctx = setup();
		void ctx.dialogs.open('cf-manager');
		const row = rowOf(dialogEl(ctx, 'cf-manager'), '=first');
		setValue(row.querySelector<HTMLInputElement>('input.xve-input')!, 'nope!!');
		clickButton(dialogEl(ctx, 'cf-manager'), 'OK');
		await tick();
		expect(ctx.toasts).toHaveLength(1);
		clickButton(dialogEl(ctx, 'cf-manager'), 'Cancel');
		const again = ctx.dialogs.open('cf-manager');
		rowOf(dialogEl(ctx, 'cf-manager'), '=first').click();
		clickButton(dialogEl(ctx, 'cf-manager'), 'Delete Rule');
		pressKey(dialogEl(ctx, 'cf-manager').querySelector('select')!, 'Escape');
		expect(await again).toBeUndefined();
		expect(order(ctx)).toEqual(['third', 'second', 'first']);
		expect(ctx.session()!.undoLabel()).toBe('Conditional formatting');
	});

	it('adds a new rule through the rule editor without applying it early', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('cf-manager');
		clickButton(dialogEl(ctx, 'cf-manager'), 'New Rule...');
		const editor = dialogEl(ctx, 'cf-rule');
		editor.querySelector<HTMLElement>('[role="option"][data-value="formula"]')!.click();
		setValue(inputByLabel(editor, 'Format values where this formula is true:'), '=TRUE');
		clickButton(editor, 'OK');
		await tick();
		expect(order(ctx)).toHaveLength(3);
		clickButton(dialogEl(ctx, 'cf-manager'), 'OK');
		await result;
		expect(order(ctx)).toEqual(['TRUE', 'third', 'second', 'first']);
	});
});
