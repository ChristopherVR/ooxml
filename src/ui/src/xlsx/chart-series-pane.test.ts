// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { buildChart } from 'ooxml-core/xlsx/ui';
import { createTestContext } from './commands/test-support';
import { chartSeriesCommand, createChartSeriesPane } from './chart-series-pane';

afterEach(() => document.body.replaceChildren());

function setup() {
	const ctx = createTestContext();
	ctx.session()!.setRangeValues(0, { row: 0, col: 0 }, [
		['Month', 'Sales'],
		['Jan', 10],
		['Feb', 20],
	]);
	const chart = buildChart(
		ctx.workbook()!,
		0,
		{ start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
		'column',
	);
	const { kind: _kind, ...model } = chart;
	const index = ctx.session()!.addChart(0, { ...model, barGapWidth: 100, barOverlap: -24 });
	ctx.selection.set({ drawing: index });
	ctx.commands.register(chartSeriesCommand());
	const pane = createChartSeriesPane(ctx);
	ctx.root.append(pane.element);
	ctx.onModelChange(pane.refresh);
	ctx.selection.onChange(pane.refresh);
	const input = (key: string) =>
		pane.element.querySelector<HTMLInputElement>(`[aria-label="${key}"]`)!;
	const change = (key: string, value: string) => {
		const control = input(key);
		control.value = value;
		control.dispatchEvent(new Event('change'));
	};
	return { ctx, pane, input, change };
}

describe('chart series pane', () => {
	it('edits live chart spacing, follows undo/redo, and saves the values', async () => {
		const { ctx, pane, input, change } = setup();
		await ctx.commands.run('chart.format-series');
		expect(pane.element.hidden).toBe(false);
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
		const { ctx, pane, input, change } = setup();
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
		expect(pane.element.querySelector<HTMLElement>('.xve-chart-series-body')!.hidden).toBe(true);
	});

	it('closes with Escape without changing selection', async () => {
		const { ctx, pane, input } = setup();
		await ctx.commands.run('chart.format-series');
		input('Gap Width').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }),
		);
		expect(pane.element.hidden).toBe(true);
		expect(ctx.selection.get().drawing).toBe(0);
	});

	it('guards a pane already open when objects become protected', async () => {
		const { ctx, pane, input, change } = setup();
		await ctx.commands.run('chart.format-series');
		ctx.workbook()!.sheets[0]!.protection = { sheet: true };
		pane.refresh();
		expect(input('Gap Width').disabled).toBe(true);
		change('Gap Width', '5');
		expect(input('Gap Width').value).toBe('100');
		expect(await ctx.commands.run('chart.format-series')).toBe(false);
	});

	it('pairs slider input with the percentage field and commits one undo step', async () => {
		const { ctx, pane, input } = setup();
		await ctx.commands.run('chart.format-series');
		const range = pane.element.querySelector<HTMLInputElement>(
			'input[type="range"][aria-label="Gap Width"]',
		)!;
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
});
