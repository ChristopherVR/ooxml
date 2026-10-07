import { expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { loadXlsx } from '../read';
import { saveXlsx } from '../write';
import {
	chartElementFill,
	chartElementFillPatch,
	chartElementGradientPatch,
	chartElementTransparencyPatch,
	chartElementTransparency,
	chartWithElementFills,
} from './chart-element-fill';
import type { ChartObject } from '../model';

function setup() {
	const book = createWorkbook(),
		session = createEditSession(book);
	session.addChart(0, {
		chartType: 'column',
		showLegend: false,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		series: [
			{
				categories: ['A', 'B'],
				values: [10, 20],
				pointColors: { 0: { kind: 'srgb', value: 'FF0000', transforms: [] } },
			},
		],
		formatting: {
			sourceXml: 'retained',
			entries: {
				chartArea: { sourceXml: 'keep', bold: true },
				categoryAxis: { sourceXml: '', labelsVisible: false },
			},
		},
	});
	return { book, session, chart: book.sheets[0]!.drawings[0] as ChartObject };
}

it('edits background fills atomically, preserving series overrides and other formatting through undo and export', async () => {
	const { book, session, chart } = setup();
	const before = structuredClone(chart);
	const patch = chartElementFillPatch(chart, 'chartArea', { theme: 5, tint: 0.25 })!;
	session.updateChart(0, 0, patch);
	expect(chart.series).toEqual(before.series);
	expect(chart.formatting?.entries.chartArea?.bold).toBe(true);
	expect(chart.formatting?.entries.categoryAxis).toEqual(before.formatting?.entries.categoryAxis);
	expect(chart).not.toHaveProperty('elementFills');
	session.updateChart(0, 0, chartElementTransparencyPatch(chart, 'chartArea', 37)!);
	expect(chartElementTransparency(chart, 'chartArea')).toBe(37);
	expect(chartElementTransparencyPatch(chart, 'chartArea', 37)).toBeUndefined();
	const back = (await loadXlsx(await saveXlsx(book))).sheets[0]!.drawings[0] as ChartObject;
	expect(chartElementTransparency(back, 'chartArea')).toBe(37);
	session.undo();
	expect(chartElementTransparency(book.sheets[0]!.drawings[0] as ChartObject, 'chartArea')).toBe(0);
	session.undo();
	expect(book.sheets[0]!.drawings[0]).toEqual(before);
});

it.each(['chartArea', 'plotArea'] as const)(
	'shares geometry and stop edits for %s without mutating preview models',
	async (part) => {
		const { book, session, chart } = setup();
		const before = structuredClone(chart);
		const created = chartElementGradientPatch(chart, part, { kind: 'create' })!;
		const preview = chartWithElementFills(chart, created.patch.elementFills!);
		expect(chart).toEqual(before);
		expect(chartElementFill(preview, part).kind).toBe('gradient');
		session.updateChart(0, 0, created.patch);
		for (const type of ['rect', 'circle', 'shape', 'linear'] as const) {
			const next = chartElementGradientPatch(chart, part, { kind: 'geometry', type });
			if (next) session.updateChart(0, 0, next.patch);
		}
		session.updateChart(
			0,
			0,
			chartElementGradientPatch(chart, part, {
				kind: 'stop',
				index: 0,
				color: { rgb: 'FF0000' },
				transparency: 37,
				position: 25,
				brightness: -20,
			})!.patch,
		);
		const saved = (await loadXlsx(await saveXlsx(book))).sheets[0]!.drawings[0] as ChartObject;
		const fill = chartElementFill(saved, part);
		expect(fill).toMatchObject({
			kind: 'gradient',
			angle: 90,
			scaled: true,
			stops: [{ position: 25 }, { position: 100 }],
		});
		expect(saved.series).toMatchObject(chart.series);
		const gradient = chartElementFill(chart, part);
		if (gradient.kind !== 'gradient') throw new Error('Expected gradient');
		expect(
			chartElementGradientPatch(chart, part, { kind: 'geometry', type: 'linear' }),
		).toBeUndefined();
	},
);

it('rejects unsupported fill targets before changing the workbook or history', () => {
	const { session, chart } = setup(),
		before = structuredClone(chart),
		label = session.undoLabel();
	expect(() =>
		session.updateChart(0, 0, { elementFills: { title: { kind: 'none' } } } as never),
	).toThrow('Unsupported chart fill target');
	expect(() => chartElementTransparencyPatch(chart, 'chartArea', NaN)).toThrow();
	expect(chart).toEqual(before);
	expect(session.undoLabel()).toBe(label);
});
