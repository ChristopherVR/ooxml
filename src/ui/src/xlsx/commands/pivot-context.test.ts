import { afterEach, expect, it } from 'vitest';
import { commandTabs, installCommands } from './index';
import { clickButton, createTestContext, inputByLabel, waitForDialog } from './test-support';

afterEach(() => document.body.replaceChildren());

it('shows both contextual tabs for a PivotTable cell and drives undoable ribbon commands', async () => {
	const ctx = createTestContext();
	const stop = installCommands(ctx);
	try {
		const session = ctx.session()!;
		const data = [
			['Group', 'Item', 'Amount'],
			['North', 'A', '10'],
			['South', 'B', '3'],
		];
		data.forEach((row, r) => row.forEach((cell, c) => session.setCellInput(0, r, c, cell)));
		session.createPivotTable(
			0,
			{ row: 0, col: 4 },
			{
				source: 'A1:C3',
				rowField: 'Group',
				secondRowField: 'Item',
				valueField: 'Amount',
			},
		);
		const tabs = commandTabs().filter(({ id }) => id === 'pivot-analyze' || id === 'pivot-design');
		expect(tabs.map((tab) => tab.contextual?.(ctx))).toEqual([false, false]);
		ctx.selection.set({ active: { row: 1, col: 4 } });
		expect(tabs.map((tab) => tab.contextual?.(ctx))).toEqual([true, true]);
		const report = () => session.workbook.sheets[0]!.pivotTables![0]!;
		const rename = ctx.commands.run('pivot.name');
		const dialog = await waitForDialog(ctx, 'pivot-name');
		inputByLabel<HTMLInputElement>(dialog, 'PivotTable Name').value = 'Sales';
		clickButton(dialog, 'OK');
		expect(await rename).toBe(true);
		expect(report().name).toBe('Sales');
		await ctx.commands.run('pivot.layout-outline');
		await ctx.commands.run('pivot.showRowStripes');
		expect(report()).toMatchObject({
			layout: 'outline',
			display: { style: { showRowStripes: true } },
		});
		session.undo();
		expect(report().display?.style?.showRowStripes).toBeUndefined();
		await ctx.commands.run('pivot.select');
		expect(ctx.selection.get().ranges).toEqual([report().location]);
		await ctx.commands.run('pivot.collapse-field');
		expect(report().display?.collapsed).toBe(true);
		await ctx.commands.run('pivot.expand-field');
		expect(report().display?.collapsed).toBe(false);
	} finally {
		stop();
	}
});
