// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { buildChart } from 'ooxml-core/xlsx/ui';
import { createTestContext } from './commands/test-support';
import { chartSeriesCommand, createChartSeriesPane } from './chart-series-pane';

afterEach(() => document.body.replaceChildren());

it('keeps the chosen series through fill edits and undo and guards stale color menus', async () => {
	const ctx = createTestContext();
	ctx.session()!.setRangeValues(0, { row: 0, col: 0 }, [
		['Month', 'Sales', 'Cost'],
		['Jan', 10, 5],
		['Feb', 20, 8],
	]);
	const { kind: _kind, ...chart } = buildChart(
		ctx.workbook()!,
		0,
		{ start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
		'column',
	);
	ctx.selection.set({ drawing: ctx.session()!.addChart(0, chart) });
	ctx.commands.register(chartSeriesCommand());
	const pane = createChartSeriesPane(ctx);
	ctx.root.append(pane.element);
	ctx.onModelChange(pane.refresh);
	await ctx.commands.run('chart.format-series', 1);
	const series = pane.element.querySelector<HTMLSelectElement>('[aria-label="Series"]')!;
	const fill = pane.element.querySelector<HTMLSelectElement>('[aria-label="Fill"]')!;
	const color = pane.element.querySelector<HTMLButtonElement>('[aria-label="Color"]')!;
	expect(series.value).toBe('1');
	expect(series.selectedOptions[0]!.textContent).toBe('Cost');
	color.click();
	series.value = '0';
	series.dispatchEvent(new Event('change'));
	ctx.root.querySelector<HTMLButtonElement>('[role="menuitem"][aria-label="Red"]')!.click();
	const drawing = ctx.workbook()!.sheets[0]!.drawings[0]!;
	if (drawing.kind !== 'chart') throw new Error('Expected chart');
	expect(drawing.series.every((item) => !item.drawingColor)).toBe(true);
	series.value = '1';
	series.dispatchEvent(new Event('change'));
	color.click();
	ctx.root.querySelector<HTMLButtonElement>('[role="menuitem"][aria-label="Red"]')!.click();
	const transparency = pane.element.querySelector<HTMLInputElement>(
		'input[type="number"][aria-label="Transparency"]',
	)!;
	transparency.value = '37';
	transparency.dispatchEvent(new Event('change'));
	expect(transparency.value).toBe('37');
	for (const value of ['101', '-1', '1.5', '']) {
		transparency.value = value;
		transparency.dispatchEvent(new Event('change'));
		expect(transparency.value).toBe('37');
	}
	expect(ctx.toasts).toHaveLength(4);
	color.click();
	ctx.root.querySelector<HTMLButtonElement>('[role="menuitem"][aria-label="Blue"]')!.click();
	expect(transparency.value).toBe('37');
	fill.value = 'none';
	fill.dispatchEvent(new Event('change'));
	ctx.session()!.undo();
	expect(series.value).toBe('1');
	expect(fill.value).toBe('solid');
	const edited = ctx.workbook()!.sheets[0]!.drawings[0]!;
	if (edited.kind !== 'chart') throw new Error('Expected chart');
	expect(edited.series[0]!.drawingColor).toBeUndefined();
	expect(edited.series[1]!.drawingColor).toEqual({
		kind: 'srgb',
		value: '0070C0',
		transforms: [{ name: 'alpha', value: '63000' }],
	});
	ctx.setReadOnly(true);
	pane.refresh();
	expect(color.disabled).toBe(true);
	expect(transparency.disabled).toBe(true);
	transparency.value = '23';
	transparency.dispatchEvent(new Event('change'));
	expect(transparency.value).toBe('37');
	fill.value = 'none';
	fill.dispatchEvent(new Event('change'));
	expect(edited.series[1]!.fill).toBeUndefined();
});
