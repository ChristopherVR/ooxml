// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { chartElementFill, chartElementTransparency, loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { buildChart } from 'ooxml-core/xlsx/ui';
import { createTestContext } from './commands/test-support';
import { chartAreaCommand, chartSeriesCommand, createChartFormatPane } from './chart-format-pane';

afterEach(() => document.body.replaceChildren());

type Pane = ReturnType<typeof createChartFormatPane>;
const shadow = <T extends Element>(pane: Pane, selector: string) =>
	pane.element.shadowRoot!.querySelector<T>(selector)!;
const title = (pane: Pane) => shadow(pane, 'h3').textContent;

function mount(ctx: ReturnType<typeof createTestContext>) {
	ctx.commands.register(chartAreaCommand());
	ctx.commands.register(chartSeriesCommand());
	const pane = createChartFormatPane(ctx);
	ctx.root.append(pane.element);
	ctx.onModelChange(pane.refresh);
	ctx.selection.onChange(pane.refresh);
	return pane;
}

function columnChart() {
	const ctx = createTestContext();
	ctx.session()!.setRangeValues(0, { row: 0, col: 0 }, [
		['Month', 'Sales'],
		['Jan', 10],
		['Feb', 20],
	]);
	const { kind: _kind, ...model } = buildChart(
		ctx.workbook()!,
		0,
		{ start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
		'column',
	);
	const index = ctx.session()!.addChart(0, { ...model, barGapWidth: 100, barOverlap: -24 });
	ctx.selection.set({ drawing: index });
	const pane = mount(ctx);
	const input = (key: string) =>
		shadow<HTMLInputElement>(pane, `input[type="number"][aria-label="${key}"]`);
	const change = (key: string, value: string) => {
		const control = input(key);
		control.value = value;
		control.dispatchEvent(new Event('change'));
	};
	return { ctx, pane, input, change };
}

describe('Excel chart Format pane', () => {
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
		const pane = mount(ctx);
		const picker = () => shadow<HTMLSelectElement>(pane, '[aria-label="Chart element"]');
		for (const [part, label] of [
			['title', 'Format Chart Title'],
			['legend', 'Format Legend'],
		] as const) {
			await ctx.commands.run('chart.format-area', part);
			expect(title(pane)).toBe(label);
			const fill = pane.element.querySelector<HTMLSelectElement>('[aria-label="Fill"]')!;
			fill.value = 'gradient';
			fill.dispatchEvent(new Event('change'));
			const chart = ctx.workbook()!.sheets[0]!.drawings[index]!;
			if (chart.kind !== 'chart') throw new Error('Missing chart');
			expect(chartElementFill(chart, part).kind).toBe('gradient');
		}
		ctx.session()!.updateChart(0, index, { showLegend: false });
		expect(picker().value).toBe('chartArea');
		expect([...picker().options].map((option) => option.value)).not.toContain('legend');
		ctx.session()!.updateChart(0, index, { title: '' });
		expect([...picker().options].map((option) => option.value)).not.toContain('title');
		await ctx.commands.run('chart.format-area', 'title');
		expect(picker().value).toBe('chartArea');
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
		const pane = mount(ctx);
		await ctx.commands.run('chart.format-area');
		const input = (label: string) =>
			pane.element.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;
		const change = (control: HTMLInputElement | HTMLSelectElement, value: string) => {
			control.value = value;
			control.dispatchEvent(new Event('change'));
		};
		expect(title(pane)).toBe('Format Chart Area');
		// A line chart's series are not formatted here: only the backgrounds are listed.
		expect(
			[...shadow<HTMLSelectElement>(pane, '[aria-label="Chart element"]').options].map(
				(option) => option.value,
			),
		).toEqual(['chartArea', 'plotArea']);
		change(input('Fill'), 'gradient');
		change(input('Type'), 'circle');
		const chart = ctx.workbook()!.sheets[0]!.drawings[0]!;
		if (chart.kind !== 'chart') throw new Error('Missing chart');
		expect(chartElementFill(chart, 'chartArea')).toMatchObject({
			kind: 'gradient',
			path: 'circle',
		});
		change(shadow<HTMLSelectElement>(pane, '[aria-label="Chart element"]'), 'plotArea');
		expect(title(pane)).toBe('Format Plot Area');
		change(input('Fill'), 'solid');
		change(input('Transparency'), '37');
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
		expect(pane.element.disabled).toBe(true);
		change(input('Fill'), 'none');
		expect(chartElementTransparency(chart, 'plotArea')).toBe(37);
		pane.close();
		expect(pane.element.hidden).toBe(true);
	});

	it('edits live chart spacing, follows undo/redo, and saves the values', async () => {
		const { ctx, pane, input, change } = columnChart();
		await ctx.commands.run('chart.format-series');
		expect(pane.element.hidden).toBe(false);
		expect(title(pane)).toBe('Format Data Series');
		expect(input('Gap Width').value).toBe('100');
		expect(input('Series Overlap').value).toBe('-24');
		change('Gap Width', '5');
		change('Series Overlap', '23');
		ctx.session()!.undo();
		expect(input('Series Overlap').value).toBe('-24');
		ctx.session()!.redo();
		expect(input('Series Overlap').value).toBe('23');
		const saved = await loadXlsx(await saveXlsx(ctx.workbook()!));
		expect(saved.sheets[0]!.drawings[0]).toMatchObject({ barGapWidth: 5, barOverlap: 23 });
	});

	it('rejects invalid values and refreshes on selection and read-only changes', async () => {
		const { ctx, pane, input, change } = columnChart();
		await ctx.commands.run('chart.format-series');
		for (const value of ['', '501', '-1', '1.5']) change('Gap Width', value);
		expect(input('Gap Width').value).toBe('100');
		expect(ctx.toasts).toHaveLength(4);
		ctx.setReadOnly(true);
		pane.refresh();
		expect(input('Gap Width').disabled).toBe(true);
		change('Gap Width', '5');
		expect(input('Gap Width').value).toBe('100');
		ctx.select('A1');
		expect(pane.element.hidden).toBe(true);
	});

	it('closes with Escape without changing selection', async () => {
		const { ctx, pane, input } = columnChart();
		await ctx.commands.run('chart.format-series');
		input('Gap Width').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, composed: true }),
		);
		expect(pane.element.hidden).toBe(true);
		expect(ctx.selection.get().drawing).toBe(0);
	});

	it('guards a pane already open when objects become protected', async () => {
		const { ctx, pane, input, change } = columnChart();
		await ctx.commands.run('chart.format-series');
		ctx.workbook()!.sheets[0]!.protection = { sheet: true };
		pane.refresh();
		expect(input('Gap Width').disabled).toBe(true);
		change('Gap Width', '5');
		expect(input('Gap Width').value).toBe('100');
		expect(await ctx.commands.run('chart.format-series')).toBe(false);
	});

	it('pairs slider input with the percentage field and commits one undo step', async () => {
		const { ctx, pane, input } = columnChart();
		await ctx.commands.run('chart.format-series');
		const range = shadow<HTMLInputElement>(pane, 'input[type="range"][aria-label="Gap Width"]');
		const original = ctx.workbook()!.sheets[0]!.drawings[0];
		for (const value of ['250', '300', '500']) {
			range.value = value;
			range.dispatchEvent(new Event('input'));
			expect(input('Gap Width').value).toBe(value);
			expect(range.getAttribute('aria-valuetext')).toBe(`${value}%`);
		}
		expect(original).toMatchObject({ barGapWidth: 100 });
		range.dispatchEvent(new Event('change'));
		expect(ctx.workbook()!.sheets[0]!.drawings[0]).toMatchObject({ barGapWidth: 500 });
		ctx.session()!.undo();
		expect(input('Gap Width').value).toBe('100');
		expect(range.value).toBe('100');
		ctx.setReadOnly(true);
		pane.refresh();
		expect(range.disabled).toBe(true);
		range.value = '5';
		range.dispatchEvent(new Event('input'));
		range.dispatchEvent(new Event('change'));
		expect(ctx.workbook()!.sheets[0]!.drawings[0]).toMatchObject({ barGapWidth: 100 });
	});

	it('moves the legend through its options', async () => {
		const ctx = createTestContext();
		const index = ctx.session()!.addChart(0, {
			chartType: 'column',
			showLegend: true,
			legendPosition: 'r',
			anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
			series: [{ categories: ['A'], values: [1] }],
		});
		ctx.selection.set({ drawing: index });
		const pane = mount(ctx);
		await ctx.commands.run('chart.format-area', 'legend');
		const bottom = shadow<HTMLElement & { checked: boolean }>(
			pane,
			'[data-chart-format-field="legend.position"] [data-chart-format-choice="b"]',
		);
		bottom.checked = true;
		bottom.dispatchEvent(new Event('change'));
		expect(ctx.workbook()!.sheets[0]!.drawings[index]).toMatchObject({ legendPosition: 'b' });
		expect(
			pane.element.shadowRoot!.querySelector('[data-chart-format-field="legend.noOverlap"]'),
		).toBeNull();
	});
});
