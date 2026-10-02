// @vitest-environment jsdom
import { createWorkbook } from '@christophervr/xlsx-core';
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

function setup() {
	const ctx = createTestContext(createWorkbook());
	registerRuleDialogs(ctx);
	ctx.dialogs.register('format-cells', async () => ({
		fill: { type: 'pattern', pattern: 'solid', bgColor: { rgb: 'FFFFFF00' } },
	}));
	ctx.select('B2:B9');
	return ctx;
}
const formats = (ctx: ReturnType<typeof setup>) => ctx.workbook()!.sheets[0]!.conditionalFormats;
const chooseType = (dialog: HTMLElement, value: string) =>
	dialog.querySelector<HTMLElement>(`[role="option"][data-value="${value}"]`)!.click();

describe('new formatting rule dialog', () => {
	it('defaults to a 2-color scale', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('cf-rule');
		clickButton(dialogEl(ctx, 'cf-rule'), 'OK');
		await result;
		expect(formats(ctx)[0]!.rules[0]).toMatchObject({
			type: 'colorScale',
			thresholds: [{ type: 'min' }, { type: 'max' }],
		});
	});

	it('builds a 3-color scale with a percentile midpoint', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('cf-rule');
		const dialog = dialogEl(ctx, 'cf-rule');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Format Style'), '3color');
		clickButton(dialog, 'OK');
		await result;
		const rule = formats(ctx)[0]!.rules[0]!;
		expect(rule).toMatchObject({
			type: 'colorScale',
			thresholds: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
		});
		expect(rule.type === 'colorScale' && rule.colors).toHaveLength(3);
	});

	it('builds data bars and icon sets', async () => {
		const ctx = setup();
		let result = ctx.dialogs.open('cf-rule');
		let dialog = dialogEl(ctx, 'cf-rule');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Format Style'), 'dataBar');
		clickButton(dialog, 'OK');
		await result;
		result = ctx.dialogs.open('cf-rule');
		dialog = dialogEl(ctx, 'cf-rule');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Format Style'), 'iconSet');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Icon Style'), '4Arrows');
		clickButton(dialog, 'OK');
		await result;
		const [bars, icons] = formats(ctx).map((f) => f.rules[0]!);
		expect(bars).toMatchObject({ type: 'dataBar', min: { type: 'min' }, max: { type: 'max' } });
		expect(icons).toMatchObject({ type: 'iconSet', iconSet: '4Arrows' });
		expect(icons?.type === 'iconSet' && icons.thresholds).toHaveLength(4);
	});

	it('builds a cell value rule with the Format... style', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('cf-rule');
		const dialog = dialogEl(ctx, 'cf-rule');
		chooseType(dialog, 'contains');
		setValue(inputByLabel(dialog, 'Value'), '1');
		setValue(inputByLabel(dialog, 'and'), '=C1');
		clickButton(dialog, 'Format...');
		await tick();
		clickButton(dialog, 'OK');
		await result;
		expect(formats(ctx)[0]!.rules[0]).toMatchObject({
			type: 'cellIs',
			operator: 'between',
			formulas: ['1', 'C1'],
			style: { fill: { bgColor: { rgb: 'FFFFFF00' } } },
		});
	});

	it('covers text, blanks, top, average, unique and formula rules', async () => {
		const ctx = setup();
		const add = async (fill: (d: HTMLElement) => void) => {
			const result = ctx.dialogs.open('cf-rule');
			const dialog = dialogEl(ctx, 'cf-rule');
			fill(dialog);
			clickButton(dialog, 'OK');
			await result;
			return formats(ctx).at(-1)!.rules[0];
		};
		expect(
			await add((d) => {
				chooseType(d, 'contains');
				setValue(inputByLabel<HTMLSelectElement>(d, 'Cells with'), 'text');
				setValue(inputByLabel<HTMLSelectElement>(d, 'Text operator'), 'endsWith');
				setValue(inputByLabel(d, 'Text'), 'x');
			}),
		).toMatchObject({ type: 'endsWith', text: 'x' });
		expect(
			await add((d) => {
				chooseType(d, 'contains');
				setValue(inputByLabel<HTMLSelectElement>(d, 'Cells with'), 'notContainsErrors');
			}),
		).toMatchObject({ type: 'notContainsErrors' });
		expect(
			await add((d) => {
				chooseType(d, 'top');
				setValue(inputByLabel<HTMLSelectElement>(d, 'Rank'), 'bottom');
				setValue(inputByLabel(d, 'Count'), '4');
			}),
		).toMatchObject({ type: 'top10', rank: 4, bottom: true });
		expect(
			await add((d) => {
				chooseType(d, 'average');
				setValue(inputByLabel<HTMLSelectElement>(d, 'Format values that are:'), 'equalBelow');
			}),
		).toMatchObject({ type: 'aboveAverage', below: true, equalAverage: true });
		expect(
			await add((d) => {
				chooseType(d, 'unique');
				setValue(inputByLabel<HTMLSelectElement>(d, 'Format all:'), 'uniqueValues');
			}),
		).toMatchObject({ type: 'uniqueValues' });
		expect(
			await add((d) => {
				chooseType(d, 'formula');
				setValue(inputByLabel(d, 'Format values where this formula is true:'), '=MOD(ROW(),2)=0');
			}),
		).toMatchObject({ type: 'expression', formula: 'MOD(ROW(),2)=0' });
	});

	it('edits an existing rule in place, keeping its priority', async () => {
		const ctx = setup();
		ctx.session()!.addConditionalFormat(0, {
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 3, col: 0 } }],
			rules: [{ type: 'expression', formula: 'A1>1', style: {}, priority: 1 }],
		});
		const result = ctx.dialogs.open('cf-rule', { format: 0, rule: 0 });
		const dialog = dialogEl(ctx, 'cf-rule');
		const input = inputByLabel(dialog, 'Format values where this formula is true:');
		expect(input.value).toBe('=A1>1');
		setValue(input, '=A1>2');
		clickButton(dialog, 'OK');
		await result;
		expect(formats(ctx)).toHaveLength(1);
		expect(formats(ctx)[0]!.rules[0]).toMatchObject({ formula: 'A1>2', priority: 1 });
	});

	it('Escape and Cancel add nothing; an empty formula keeps the dialog open', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('cf-rule');
		pressKey(dialogEl(ctx, 'cf-rule').querySelector('[role="listbox"]')!, 'Escape');
		expect(await result).toBeUndefined();
		void ctx.dialogs.open('cf-rule');
		const dialog = dialogEl(ctx, 'cf-rule');
		chooseType(dialog, 'formula');
		clickButton(dialog, 'OK');
		await tick();
		expect(ctx.toasts).toHaveLength(1);
		clickButton(dialog, 'Cancel');
		expect(formats(ctx)).toEqual([]);
	});
});
