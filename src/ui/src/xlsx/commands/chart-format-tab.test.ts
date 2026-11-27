// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { chartTargetShape } from 'ooxml-core/chart';
import {
	anchorToPixelBox,
	chartSpaceOf,
	createGridMetrics,
	type ChartObject,
} from 'ooxml-core/xlsx';
import { buildChart } from 'ooxml-core/xlsx/ui';
import { chartElementOptions, chartFormatTabCommands } from './chart-format-tab';
import { commandTabs } from './index';
import { clickButton, createTestContext, inputByLabel, waitForDialog } from './test-support';

afterEach(() => document.body.replaceChildren());

function chartContext() {
	const ctx = createTestContext();
	const session = ctx.session()!;
	session.setRangeValues(0, { row: 0, col: 0 }, [
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
	const index = session.addChart(0, model);
	ctx.commands.registerAll(chartFormatTabCommands());
	const opened: unknown[] = [];
	ctx.dialogs.register('format-chart-selection', async (_ctx, arg) => void opened.push(arg));
	ctx.selection.set({ drawing: index });
	const chart = () => ctx.workbook()!.sheets[0]!.drawings[index] as ChartObject;
	const space = () => chartSpaceOf(ctx.workbook()!, chart());
	return { ctx, session, index, opened, space, chart };
}

describe('Excel chart Format tab', () => {
	it('is a contextual tab with the groups of Excel', () => {
		const { ctx } = chartContext();
		const tab = commandTabs().find(({ id }) => id === 'chart-format')!;
		expect(tab.contextual?.(ctx)).toBe(true);
		expect(tab.groups.map((group) => group.label)).toEqual([
			'Current Selection',
			'Insert Shapes',
			'Shape Styles',
			'WordArt Styles',
			'Arrange',
			'Size',
		]);
		ctx.selection.set({ drawing: undefined });
		expect(tab.contextual?.(ctx)).toBe(false);
	});

	it('formats the element picked in Chart Elements and opens the pane on it', async () => {
		const { ctx, opened, space } = chartContext();
		const options = chartElementOptions(ctx);
		expect(options.map((option) => option.label)).toContain('Chart Area');
		expect(options.some((option) => option.label === 'Vertical (Value) Axis')).toBe(true);
		const series = options.find((option) => option.value === 'series:0')!;
		expect(series.label).toBe('Series "Sales"');
		expect(ctx.commands.get('chart.element')?.value?.(ctx)).toBe('chartArea');
		await ctx.commands.run('chart.element', 'series:0');
		expect(ctx.commands.get('chart.element')?.value?.(ctx)).toBe('series:0');

		await ctx.commands.run('chart.shape-fill', { rgb: 'FF0000' });
		expect(chartTargetShape(space(), { kind: 'series', seriesIndex: 0 })?.fill).toMatchObject({
			kind: 'solid',
			color: { value: 'FF0000' },
		});
		await ctx.commands.run('chart.shape-outline', undefined);
		expect(chartTargetShape(space(), { kind: 'series', seriesIndex: 0 })?.line?.fill).toEqual({
			kind: 'none',
		});
		await ctx.commands.run('chart.shape-effects', true);
		expect(chartTargetShape(space(), { kind: 'series', seriesIndex: 0 })?.effectsXml).toContain(
			'outerShdw',
		);
		await ctx.commands.run('chart.format-selection');
		expect(opened).toEqual(['series:0']);

		await ctx.commands.run('chart.reset-style');
		const reset = chartTargetShape(space(), { kind: 'series', seriesIndex: 0 });
		expect(reset?.fill).toBeUndefined();
		expect(reset?.effectsXml ?? '').not.toContain('outerShdw');
		expect(ctx.session()!.undo()).toBe(true);
		expect(chartTargetShape(space(), { kind: 'series', seriesIndex: 0 })?.fill).toBeDefined();
	});

	it('applies a Shape Styles pick and a WordArt style to the chart area', async () => {
		const { ctx, space } = chartContext();
		await ctx.commands.run('chart.shape-style', 'preset-3-1');
		expect(chartTargetShape(space(), { kind: 'chartArea' })?.fill?.kind).toBe('solid');
		await ctx.commands.run('chart.shape-style', 'theme-0-1');
		expect(chartTargetShape(space(), { kind: 'chartArea' })?.line).toBeDefined();
		await ctx.commands.run('chart.word-art-style', '1');
		expect(JSON.stringify(space())).toContain('1769A5');
	});

	it('resizes the chart from Size', async () => {
		const { ctx, chart } = chartContext();
		const run = ctx.commands.run('chart.size');
		const dialog = await waitForDialog(ctx, 'chart-size');
		inputByLabel<HTMLInputElement>(dialog, 'Width (px)').value = '300';
		inputByLabel<HTMLInputElement>(dialog, 'Height (px)').value = '150';
		clickButton(dialog, 'OK');
		expect(await run).toBe(true);
		const sheet = ctx.workbook()!.sheets[0]!;
		const box = anchorToPixelBox(sheet, chart().anchor, createGridMetrics(sheet, { zoom: 100 }));
		expect([Math.round(box.w), Math.round(box.h)]).toEqual([300, 150]);
	});
});
