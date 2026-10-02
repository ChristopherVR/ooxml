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
import { quickRule } from './cf-quick.js';
import { registerRuleDialogs } from './register-rules.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook());
	registerRuleDialogs(ctx);
	ctx.select('A1:A5');
	return ctx;
}
const sheet = (ctx: ReturnType<typeof setup>) => ctx.workbook()!.sheets[0]!;

describe('quick rule dialog', () => {
	it('adds an A Date Occurring rule for the chosen period', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('cf-quick', { kind: 'timePeriod' });
		const dialog = dialogEl(ctx, 'cf-quick');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Date'), 'last7Days');
		clickButton(dialog, 'OK');
		await result;
		expect(sheet(ctx).conditionalFormats[0]?.rules[0]).toMatchObject({
			type: 'timePeriod',
			timePeriod: 'last7Days',
		});
	});

	it('adds a greater-than rule with the chosen format', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('cf-quick', { kind: 'greaterThan' });
		const dialog = dialogEl(ctx, 'cf-quick');
		setValue(inputByLabel(dialog, 'Value'), '5');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'with'), 'green');
		clickButton(dialog, 'OK');
		await result;
		const cf = sheet(ctx).conditionalFormats[0]!;
		expect(cf.ranges).toEqual([{ start: { row: 0, col: 0 }, end: { row: 4, col: 0 } }]);
		expect(cf.rules[0]).toMatchObject({ type: 'cellIs', operator: 'greaterThan', formulas: ['5'] });
		expect(cf.rules[0]).toHaveProperty('style.fill.bgColor.rgb', 'FFC6EFCE');
	});

	it('quotes text, strips = and builds the other rule kinds', () => {
		const style = {};
		expect(quickRule('between', ['=B1', 'abc'], style, {})).toMatchObject({
			formulas: ['B1', '"abc"'],
		});
		expect(quickRule('containsText', ['x'], style, {})).toMatchObject({
			type: 'containsText',
			text: 'x',
		});
		expect(quickRule('duplicateValues', [], style, { unique: true }).type).toBe('uniqueValues');
		expect(quickRule('bottom10Percent', [], style, { rank: 5 })).toMatchObject({
			type: 'top10',
			rank: 5,
			bottom: true,
			percent: true,
		});
		expect(quickRule('belowAverage', [], style, {})).toMatchObject({
			type: 'aboveAverage',
			below: true,
		});
	});

	it('uses a custom format from Format Cells', async () => {
		const ctx = setup();
		ctx.dialogs.register('format-cells', async () => ({ font: { bold: true } }));
		const result = ctx.dialogs.open('cf-quick', { kind: 'top10' });
		const dialog = dialogEl(ctx, 'cf-quick');
		setValue(inputByLabel(dialog, 'Top'), '3');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'with'), 'custom');
		await tick();
		clickButton(dialog, 'OK');
		await result;
		expect(sheet(ctx).conditionalFormats[0]!.rules[0]).toMatchObject({
			type: 'top10',
			rank: 3,
			style: { font: { bold: true } },
		});
	});

	it('requires a value; Cancel and Escape add nothing', async () => {
		const ctx = setup();
		void ctx.dialogs.open('cf-quick', { kind: 'lessThan' });
		clickButton(dialogEl(ctx, 'cf-quick'), 'OK');
		await tick();
		expect(ctx.toasts).toHaveLength(1);
		clickButton(dialogEl(ctx, 'cf-quick'), 'Cancel');
		const again = ctx.dialogs.open('cf-quick', { kind: 'equal' });
		pressKey(inputByLabel(dialogEl(ctx, 'cf-quick'), 'Value'), 'Escape');
		expect(await again).toBeUndefined();
		expect(sheet(ctx).conditionalFormats).toEqual([]);
	});
});
