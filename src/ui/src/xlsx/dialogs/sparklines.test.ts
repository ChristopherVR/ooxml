import { afterEach, describe, expect, it } from 'vitest';
import { createWorkbook, parseRange, sheetSparklineGroups } from 'ooxml-core/xlsx';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	setValue,
} from '../commands/test-support';
import { registerToolDialogs } from './register-tools';
import { sparklineCommands } from '../commands/sparklines';

afterEach(() => {
	document.body.innerHTML = '';
});

describe('sparkline dialogs', () => {
	it('creates a group from data and location ranges', async () => {
		const ctx = createTestContext(createWorkbook());
		registerToolDialogs(ctx);
		ctx.select('D1:D2');
		const done = ctx.dialogs.open('create-sparklines', { type: 'column' });
		const box = dialogEl(ctx, 'create-sparklines');
		expect(inputByLabel(box, 'Location Range:').value).toBe('D1:D2');
		setValue(inputByLabel(box, 'Data Range:'), 'A1:C2');
		clickButton(box, 'OK');
		expect(await done).toBe(true);
		expect(sheetSparklineGroups(ctx.workbook()!.sheets[0]!)[0]).toMatchObject({
			type: 'column',
			sparklines: [{ host: parseRange('D1') }, { host: parseRange('D2') }],
		});
	});
	it('applies group settings and clears them through commands with undo', async () => {
		const ctx = createTestContext(createWorkbook());
		const session = ctx.session()!;
		session.createSparklineGroup(0, 'A1:C2', parseRange('D1:D2')!);
		ctx.select('D1');
		registerToolDialogs(ctx);
		const done = ctx.dialogs.open('sparkline-settings');
		const box = dialogEl(ctx, 'sparkline-settings');
		setValue(inputByLabel<HTMLSelectElement>(box, 'Sparkline Type:'), 'column');
		inputByLabel<HTMLInputElement>(box, 'High Point').click();
		setValue(inputByLabel<HTMLSelectElement>(box, 'Minimum Axis:'), 'custom');
		setValue(inputByLabel(box, 'Minimum Value:'), '-10');
		clickButton(box, 'OK');
		expect(await done).toBe(true);
		expect(sheetSparklineGroups(ctx.workbook()!.sheets[0]!)[0]).toMatchObject({
			type: 'column',
			high: true,
			manualMin: -10,
		});
		ctx.commands.registerAll(sparklineCommands());
		ctx.commands.run('sparkline.clear-selected');
		expect(sheetSparklineGroups(ctx.workbook()!.sheets[0]!)[0]?.sparklines).toHaveLength(1);
		session.undo();
		expect(sheetSparklineGroups(ctx.workbook()!.sheets[0]!)[0]?.sparklines).toHaveLength(2);
	});
	it('disables editing commands when the sheet is protected', () => {
		const wb = createWorkbook();
		const ctx = createTestContext(wb);
		wb.sheets[0]!.protection = { sheet: true };
		for (const command of sparklineCommands()) expect(command.enabled?.(ctx)).toBe(false);
	});
	it('shows mixed settings and applies only changed properties across selected groups', async () => {
		const ctx = createTestContext(createWorkbook());
		const session = ctx.session()!;
		session.createSparklineGroup(0, 'A1:C1', parseRange('D1')!, { type: 'line' });
		session.createSparklineGroup(0, 'A2:C2', parseRange('D2')!, {
			type: 'column',
			high: true,
			lineWeight: 2,
			minAxisType: 'custom',
			manualMin: -8,
			maxAxisType: 'custom',
			manualMax: 12,
			colorSeries: { rgb: 'FF0000' },
		});
		ctx.select('D1:D2');
		registerToolDialogs(ctx);
		const sheet = ctx.workbook()!.sheets[0]!;
		const before = structuredClone(sheetSparklineGroups(sheet));
		const undo = session.undoLabel();
		const unchanged = ctx.dialogs.open('sparkline-settings');
		let box = dialogEl(ctx, 'sparkline-settings');
		expect(inputByLabel<HTMLSelectElement>(box, 'Sparkline Type:').value).toBe('');
		expect(inputByLabel<HTMLInputElement>(box, 'High Point').indeterminate).toBe(true);
		setValue(inputByLabel<HTMLSelectElement>(box, 'Sparkline Type:'), 'line');
		setValue(inputByLabel<HTMLSelectElement>(box, 'Sparkline Type:'), '');
		clickButton(box, 'OK');
		expect(await unchanged).toBe(true);
		expect(sheetSparklineGroups(sheet)).toEqual(before);
		expect(session.undoLabel()).toBe(undo);
		const changed = ctx.dialogs.open('sparkline-settings');
		box = dialogEl(ctx, 'sparkline-settings');
		setValue(inputByLabel(box, 'Sparkline Color:'), '#123456');
		clickButton(box, 'OK');
		expect(await changed).toBe(true);
		const after = sheetSparklineGroups(sheet);
		expect(after.map((group) => group.type)).toEqual(['line', 'column']);
		expect(after.map((group) => group.high)).toEqual([false, true]);
		expect(after.map((group) => group.lineWeight)).toEqual([0.75, 2]);
		expect(after[1]).toMatchObject({
			minAxisType: 'custom',
			manualMin: -8,
			maxAxisType: 'custom',
			manualMax: 12,
		});
		expect(after.map((group) => group.colorSeries)).toEqual([
			{ rgb: 'FF123456' },
			{ rgb: 'FF123456' },
		]);
	});
});
