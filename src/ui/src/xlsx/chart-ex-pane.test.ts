// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { afterEach, expect, it } from 'vitest';
import { loadXlsx, xlsxChartExDataSheet } from 'ooxml-core/xlsx';
import { createTestContext } from './commands/test-support';
import { chartAreaCommand } from './chart-format-pane';
import { createChartExPane } from './chart-ex-pane';
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
	const pane = createChartExPane(ctx);
	ctx.root.append(pane.element);
	ctx.onModelChange(pane.refresh);
	await ctx.commands.run('chart.format-area');
	expect(pane.element.hidden).toBe(false);
	const options = pane.element.querySelector('office-ui-chart-ex-options')!;
	options.dispatchEvent(
		new CustomEvent('office-chart-ex-edit', {
			detail: { action: { op: 'title', text: 'Changed' } },
		}),
	);
	const drawing = ctx.workbook()!.sheets[s]!.drawings[i]!;
	expect(drawing.kind === 'unsupported' && drawing.chartExSpace?.title?.tx?.rich?.text).toBe(
		'Changed',
	);
	pane.element.querySelector('office-ui-chart-data-sheet')!.dispatchEvent(
		new CustomEvent('office-chart-data-change', {
			detail: { edits: [{ kind: 'value', series: 0, category: 0, value: 77 }] },
		}),
	);
	expect(xlsxChartExDataSheet(book, s, i).series[0]!.values[0]).toBe(77);
	ctx.session()!.undo();
	expect(xlsxChartExDataSheet(book, s, i).series[0]!.values[0]).not.toBe(77);
	ctx.setReadOnly(true);
	pane.refresh();
	expect((options as HTMLElement & { disabled: boolean }).disabled).toBe(true);
});
