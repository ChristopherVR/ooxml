// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { afterEach, expect, it } from 'vitest';
import { loadXlsx, xlsxChartExDataSheet } from 'ooxml-core/xlsx';
import { createTestContext } from './commands/test-support';
import { chartAreaCommand } from './commands/chart-format';
import { registerChartPanes } from './chart-panes';
afterEach(() => document.body.replaceChildren());
it('opens ChartEx controls, updates cells and title, and follows undo', async () => {
	const book = await loadXlsx(
		readFileSync(
			path.join(import.meta.dirname, '../../../core/xlsx/__fixtures__/excel-chart-families.xlsx'),
		),
	);
	let selected: { s: number; i: number } | undefined;
	for (const [s, ws] of book.sheets.entries())
		for (const [i, d] of ws.drawings.entries())
			if (d.kind === 'unsupported' && d.chartExSpace?.plotArea.series[0]?.layoutId === 'waterfall')
				selected = { s, i };
	const { s, i } = selected!;
	const ctx = createTestContext(book);
	ctx.setActiveSheet(s);
	ctx.selection.set({ drawing: i });
	ctx.commands.register(chartAreaCommand());
	const workspace = document.createElement('div');
	const smartArt = document.createElement('div');
	ctx.root.append(workspace, smartArt);
	const panes = registerChartPanes(ctx, workspace, smartArt);
	ctx.onModelChange(panes.refresh);
	expect(panes.elements()).toEqual([]);
	await ctx.commands.run('chart.format-area');
	const [pane] = panes.elements();
	expect(pane?.hidden).toBe(false);
	expect(workspace.nextElementSibling).toBe(pane);
	expect(pane?.nextElementSibling).toBe(smartArt);
	const options = pane!.querySelector('office-ui-chart-ex-options')!;
	options.dispatchEvent(
		new CustomEvent('office-chart-ex-edit', {
			detail: { action: { op: 'title', text: 'Changed' } },
		}),
	);
	const drawing = ctx.workbook()!.sheets[s]!.drawings[i]!;
	expect(drawing.kind === 'unsupported' && drawing.chartExSpace?.title?.tx?.rich?.text).toBe(
		'Changed',
	);
	pane!.querySelector('office-ui-chart-data-sheet')!.dispatchEvent(
		new CustomEvent('office-chart-data-change', {
			detail: { edits: [{ kind: 'value', series: 0, category: 0, value: 77 }] },
		}),
	);
	expect(xlsxChartExDataSheet(book, s, i).series[0]!.values[0]).toBe(77);
	ctx.session()!.undo();
	expect(xlsxChartExDataSheet(book, s, i).series[0]!.values[0]).not.toBe(77);
	await ctx.commands.run('chart.format-area');
	expect(panes.elements()).toEqual([pane]);
	ctx.setReadOnly(true);
	panes.refresh();
	expect((options as HTMLElement & { disabled: boolean }).disabled).toBe(true);
});
