// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { chartElementFill, chartElementTransparency, loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { createTestContext } from './commands/test-support';
import { chartAreaCommand, createChartAreaPane } from './chart-area-pane';
afterEach(() => document.body.replaceChildren());

it('offers present titles and legends and falls back when the target disappears', async () => {
	const ctx = createTestContext();
	const index = ctx.session()!.addChart(0, {
		chartType: 'column',
		title: 'Sales',
		showLegend: true,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		series: [],
	});
	ctx.selection.set({ drawing: index });
	ctx.commands.register(chartAreaCommand());
	const pane = createChartAreaPane(ctx, () => {});
	ctx.root.append(pane.element);
	ctx.onModelChange(pane.refresh);
	const select = pane.element.querySelector<HTMLSelectElement>('[aria-label="Chart element"]')!;
	for (const [part, label] of [
		['title', 'Format Chart Title'],
		['legend', 'Format Legend'],
	] as const) {
		await ctx.commands.run('chart.format-area', part);
		expect(pane.element.getAttribute('aria-label')).toBe(label);
		const fill = pane.element.querySelector<HTMLSelectElement>('[aria-label="Fill"]')!;
		fill.value = 'gradient';
		fill.dispatchEvent(new Event('change'));
		const chart = ctx.workbook()!.sheets[0]!.drawings[index]!;
		if (chart.kind !== 'chart') throw new Error('Missing chart');
		expect(chartElementFill(chart, part).kind).toBe('gradient');
	}
	ctx.session()!.updateChart(0, index, { showLegend: false });
	expect(select.value).toBe('chartArea');
	expect([...select.options].find((option) => option.value === 'legend')!.disabled).toBe(true);
	ctx.session()!.updateChart(0, index, { title: '' });
	expect([...select.options].find((option) => option.value === 'title')!.disabled).toBe(true);
	await ctx.commands.run('chart.format-area', 'title');
	expect(select.value).toBe('chartArea');
});

it('uses shared fill controls on both backgrounds, follows history and guards read-only changes', async () => {
	const ctx = createTestContext();
	const index = ctx.session()!.addChart(0, {
		chartType: 'line',
		showLegend: false,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		series: [{ categories: ['A', 'B'], values: [10, 20] }],
	});
	ctx.selection.set({ drawing: index });
	ctx.commands.register(chartAreaCommand());
	const pane = createChartAreaPane(ctx, () => {});
	ctx.root.append(pane.element);
	ctx.onModelChange(pane.refresh);
	await ctx.commands.run('chart.format-area');
	const input = (label: string) =>
		pane.element.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;
	const change = (label: string, value: string) => {
		const control = input(label);
		control.value = value;
		control.dispatchEvent(new Event('change'));
	};
	expect(pane.element.getAttribute('aria-label')).toBe('Format Chart Area');
	change('Fill', 'gradient');
	change('Type', 'circle');
	const chart = ctx.workbook()!.sheets[0]!.drawings[0]!;
	if (chart.kind !== 'chart') throw new Error('Missing chart');
	expect(chartElementFill(chart, 'chartArea')).toMatchObject({ kind: 'gradient', path: 'circle' });
	change('Chart element', 'plotArea');
	expect(pane.element.getAttribute('aria-label')).toBe('Format Plot Area');
	change('Fill', 'solid');
	change('Transparency', '37');
	expect(chartElementTransparency(chart, 'plotArea')).toBe(37);
	ctx.session()!.undo();
	expect(input('Transparency').value).toBe('0');
	ctx.session()!.redo();
	const saved = (await loadXlsx(await saveXlsx(ctx.workbook()!))).sheets[0]!.drawings[0]!;
	if (saved.kind !== 'chart') throw new Error('Missing exported chart');
	expect(chartElementTransparency(saved, 'plotArea')).toBe(37);
	ctx.setReadOnly(true);
	pane.refresh();
	expect(input('Fill').disabled).toBe(true);
	change('Fill', 'none');
	expect(chartElementTransparency(chart, 'plotArea')).toBe(37);
	pane.close();
	expect(pane.element.hidden).toBe(true);
});
