import { afterEach, expect, it } from 'vitest';
import { commandTabs, installCommands } from './index';
import { clickButton, createTestContext, inputByLabel, waitForDialog } from './test-support';

afterEach(() => document.body.replaceChildren());

it('preserves custom subtotal functions when switching positions', async () => {
	const ctx = createTestContext();
	const stop = installCommands(ctx);
	try {
		const session = ctx.session()!;
		[
			['Group', 'Item', 'Amount'],
			['North', 'A', '10'],
			['North', 'B', '3'],
		].forEach((row, r) => row.forEach((cell, c) => session.setCellInput(0, r, c, cell)));
		const custom = { field: 'Group', functions: ['count', 'average'] } as const;
		session.createPivotTable(
			0,
			{ row: 0, col: 4 },
			{
				source: 'A1:C3',
				rowField: 'Group',
				secondRowField: 'Item',
				valueField: 'Amount',
				fieldSubtotals: [
					{ ...custom, functions: [...custom.functions] },
					{ field: 'Item', functions: [] },
				],
			},
		);
		ctx.select('E2');
		for (const position of ['top', 'bottom']) {
			expect(await ctx.commands.run(`pivot.subtotal-${position}`)).toBe(true);
			expect(session.workbook.sheets[0]?.pivotTables?.[0]).toMatchObject({
				subtotalPosition: position,
				fieldSubtotals: [custom],
			});
		}
	} finally {
		stop();
	}
});

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

		// Shared PivotTables stay editable; only one over the room's size limit waits for sharing to end.
		ctx.collaborationActive = () => true;
		expect(await ctx.commands.run('pivot.showRowStripes')).toBe(true);
		expect(report().display?.style?.showRowStripes).toBe(true);
		expect(await ctx.commands.run('pivot.select')).toBe(true);
		const pivots = session.workbook.sheets[0]!.pivotTables!;
		const shared = report();
		pivots[0] = {
			...shared,
			records: Array.from({ length: 300_000 }, () => ['Group', 'Item', 1]),
		};
		ctx.selection.set({ active: shared.location.start });
		expect(ctx.commands.isEnabled('pivot.refresh')).toBe(false);
		expect(ctx.commands.isEnabled('insert.pivot-table')).toBe(true);
		ctx.collaborationActive = () => false;
		expect(ctx.commands.isEnabled('pivot.refresh')).toBe(true);
		pivots[0] = shared;
	} finally {
		stop();
	}
});
