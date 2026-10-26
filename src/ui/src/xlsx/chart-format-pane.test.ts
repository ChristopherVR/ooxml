// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import {
	chartElementFill,
	chartElementTransparency,
	loadXlsx,
	saveXlsx,
	chartSpaceOf,
} from 'ooxml-core/xlsx';
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
	it('applies advanced shared controls as workbook history entries and saves their chart XML', async () => {
		const { ctx, pane } = columnChart();
		await ctx.commands.run('chart.format-series');
		const change = (intent: import('ooxml-core/chart').ChartEditIntent, field: string) =>
			pane.element.dispatchEvent(
				new CustomEvent('office-chart-format-edit', {
					detail: { target: 'series:0', field, value: '', intent },
				}),
			);
		const space = () => {
			const book = ctx.workbook()!;
			const chart = book.sheets[0]!.drawings[0]!;
			if (chart.kind !== 'chart') throw new Error('Missing chart');
			return chartSpaceOf(book, chart);
		};
		change(
			{
				op: 'trendline',
				seriesIndex: 0,
				index: 0,
				edit: { type: 'linear', displayEquation: true },
			},
			'trendline.type',
		);
		change(
			{
				op: 'errorBars',
				seriesIndex: 0,
				direction: 'y',
				edit: { valueType: 'percentage', value: 5 },
			},
			'errorBars.type',
		);
		expect(space().plotArea.groups[0]!.series[0]!.errorBars?.[0]?.value).toBe(5);
		ctx.session()!.undo();
		expect(space().plotArea.groups[0]!.series[0]!.errorBars?.length ?? 0).toBe(0);
		ctx.session()!.redo();
		const bytes = await saveXlsx(ctx.workbook()!);
		const reloaded = await loadXlsx(bytes);
		const chart = reloaded.sheets[0]!.drawings[0]!;
		if (chart.kind !== 'chart') throw new Error('Missing exported chart');
		const result = chartSpaceOf(reloaded, chart).plotArea.groups[0]!.series[0]!;
		expect(result.trendlines?.[0]?.displayEquation).toBe(true);
		expect(result.errorBars?.[0]?.value).toBe(5);
	});
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
		// Every element the chart part has, in PowerPoint's picker order.
		expect(
			[...shadow<HTMLSelectElement>(pane, '[aria-label="Chart element"]').options].map(
				(option) => option.value,
			),
		).toEqual([
			'chartArea',
			'axis:0',
			'axis:1',
			'majorGridlines:1',
			'plotArea',
			'series:0',
			'dataLabels:0',
		]);
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
		const noOverlap = shadow<HTMLElement & { checked: boolean }>(
			pane,
			'[data-chart-format-field="legend.noOverlap"] office-ui-checkbox',
		);
		noOverlap.checked = false;
		noOverlap.dispatchEvent(new Event('change'));
		const chart = ctx.workbook()!.sheets[0]!.drawings[index]!;
		if (chart.kind !== 'chart') throw new Error('Missing chart');
		expect(chart.chartSpace?.legend).toMatchObject({ position: 'b', overlay: true });
	});

	it('edits axis, gridline, border and data label options through chart edit intents', async () => {
		const ctx = createTestContext();
		const index = ctx.session()!.addChart(0, {
			chartType: 'line',
			showLegend: false,
			anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
			series: [{ name: 'Sales', categories: ['A', 'B'], values: [10, 20] }],
		});
		ctx.selection.set({ drawing: index });
		const pane = mount(ctx);
		await ctx.commands.run('chart.format-area');
		const picker = shadow<HTMLSelectElement>(pane, '[aria-label="Chart element"]');
		const pick = (key: string) => {
			picker.value = key;
			picker.dispatchEvent(new Event('change'));
		};
		const field = <T extends HTMLElement>(id: string, selector = 'input') =>
			shadow<T>(pane, `[data-chart-format-field="${id}"] ${selector}`);
		const chart = () => {
			const found = ctx.workbook()!.sheets[0]!.drawings[index]!;
			if (found.kind !== 'chart') throw new Error('Missing chart');
			return found;
		};
		pick('axis:1');
		expect(title(pane)).toBe('Format Axis');
		const max = field<HTMLInputElement>('axis.max');
		max.value = '50';
		max.dispatchEvent(new Event('change'));
		expect(chart().chartSpace?.plotArea.axes[1]?.scaling.max).toBe(50);
		expect(chart().axes?.[1]?.scaling.max).toBe(50);
		pick('series:0');
		expect(title(pane)).toBe('Format Data Series');
		const width = field<HTMLInputElement>('line.width');
		width.value = '4';
		width.dispatchEvent(new Event('change'));
		expect(chart().chartSpace?.plotArea.groups[0]?.series[0]?.spPr?.line?.widthEmu).toBe(50800);
		pick('dataLabels:0');
		const value = field<HTMLElement & { checked: boolean }>(
			'labels.showValue',
			'office-ui-checkbox',
		);
		value.checked = true;
		value.dispatchEvent(new Event('change'));
		expect(chart().series[0]?.dataLabels?.series?.showValue).toBe(true);
		ctx.session()!.undo();
		expect(chart().series[0]?.dataLabels?.series?.showValue).not.toBe(true);
		ctx.session()!.redo();
		const saved = (await loadXlsx(await saveXlsx(ctx.workbook()!))).sheets[0]!.drawings[0]!;
		if (saved.kind !== 'chart') throw new Error('Missing exported chart');
		expect(saved.axes?.[1]?.scaling.max).toBe(50);
		expect(saved.chartSpace?.plotArea.groups[0]?.series[0]?.spPr?.line?.widthEmu).toBe(50800);
		expect(saved.series[0]?.dataLabels?.series?.showValue).toBe(true);
	});
});
