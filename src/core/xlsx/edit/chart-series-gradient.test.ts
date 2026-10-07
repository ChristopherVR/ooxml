import JSZip from 'jszip';
import { expect, it } from 'vitest';
import native from '../__fixtures__/excel-chart-gradient-edits.json';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import type { ChartObject } from '../model';
import { chartView } from '../layout/chart-view';
import { buildChartGradientDef } from '../../chart/gradient-definition';
import { drawingColorBrightness } from '../../diagram/drawing-color-brightness';
import {
	chartSeriesGradientPatch,
	chartGradientStopTransparency,
	type ChartGradientEdit,
} from './chart-series-gradient';

async function setup(sample = native.cases[0]!) {
	const book = createWorkbook();
	createEditSession(book).addChart(0, {
		chartType: 'column',
		series: [],
		showLegend: false,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
	});
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	zip.file('xl/charts/chart1.xml', sample.chartXml);
	return loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
}
const chartOf = (book: ReturnType<typeof createWorkbook>) =>
	book.sheets[0]!.drawings[0] as ChartObject;

it('matches Excel brightness extremes, reset, and opacity preservation through undo and save', async () => {
	const book = await setup(native.cases[3]!);
	const session = createEditSession(book);
	for (const value of [-42, 0, 100, -100, 37]) {
		const before = structuredClone(chartOf(book).series[0]!.fill);
		const result = chartSeriesGradientPatch(chartOf(book), 0, {
			kind: 'stop',
			index: 0,
			brightness: value,
		})!;
		session.updateChart(0, 0, result.patch);
		const saved = chartOf(await loadXlsx(await saveXlsx(book))).series[0]!.fill;
		const expected = chartOf(
			await setup(native.cases.find((sample) => sample.name === `brightness-${value}`)!),
		).series[0]!.fill;
		if (saved?.kind !== 'gradient' || expected?.kind !== 'gradient')
			throw new Error('Expected gradients');
		expect(saved.stops).toEqual(expected.stops);
		expect(drawingColorBrightness(saved.stops[0]!.color)).toBe(value);
		expect(chartGradientStopTransparency(saved, 0)).toBe(37);
		session.undo();
		expect(chartOf(book).series[0]!.fill).toEqual(before);
		session.redo();
	}
	for (const brightness of [-101, 101, Number.NaN])
		expect(() =>
			chartSeriesGradientPatch(chartOf(book), 0, { kind: 'stop', index: 0, brightness }),
		).toThrow(RangeError);
});

it('creates a gradient from the chosen series palette and preserves another series', () => {
	const book = createWorkbook();
	createEditSession(book).addChart(0, {
		chartType: 'column',
		showLegend: false,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		series: [
			{ categories: ['A'], values: [1] },
			{ categories: ['A'], values: [2], fill: { kind: 'none' } },
		],
	});
	const chart = chartOf(book);
	const before = structuredClone(chart);
	const result = chartSeriesGradientPatch(chart, 1, { kind: 'create' })!;
	expect(chart).toEqual(before);
	expect(result.patch.series![0]).toEqual(before.series[0]);
	expect(result.patch.series![1]!.fill).toMatchObject({
		kind: 'gradient',
		angle: 90,
		stops: [{ position: 0, color: { kind: 'scheme', value: 'accent2' } }, { position: 100 }],
	});
});

for (const sample of native.cases)
	it(`renders native ${sample.name} gradient stops in visual order without changing their identities`, async () => {
		const book = await setup(sample);
		const chart = chartOf(book);
		const fill = chart.series[0]!.fill!;
		if (fill.kind !== 'gradient') throw new Error('Expected gradient');
		expect(fill.angle).toBe(sample.angle);
		expect(fill.stops).toHaveLength(sample.stops.length);
		fill.stops.forEach((stop, index) => {
			expect(drawingColorBrightness(stop.color)).toBeCloseTo(sample.stops[index]!.brightness, 4);
			expect(stop.position).toBeCloseTo(sample.stops[index]!.position, 4);
			expect(chartGradientStopTransparency(fill, index)).toBeCloseTo(
				sample.stops[index]!.transparency,
				4,
			);
		});
		const gradient = chartView(book, 0, chart, () => []).series[0]!.gradient!;
		expect(gradient.stops.map((stop) => stop.color)).toEqual(
			sample.stops.map((stop) => {
				const rgb = stop.rgb;
				return (
					'#' +
					[rgb & 255, (rgb >>> 8) & 255, (rgb >>> 16) & 255]
						.map((byte) => byte.toString(16).padStart(2, '0'))
						.join('')
						.toUpperCase()
				);
			}),
		);
		expect(buildChartGradientDef('native', gradient).stops.map((stop) => stop.offset)).toEqual(
			fill.stops.map((stop) => stop.position / 100).sort((a, b) => a - b),
		);
		expect((await loadXlsx(await saveXlsx(book))).sheets[0]!.drawings[0]).toMatchObject({
			series: [{ fill }],
		});
	});

it('matches native angle, stop edits, insert/delete and minimum through undo and source-preserving saves', async () => {
	const book = await setup();
	const session = createEditSession(book);
	const apply = (edit: ChartGradientEdit) => {
		const result = chartSeriesGradientPatch(chartOf(book), 0, edit)!;
		session.updateChart(0, 0, result.patch);
		return result.stopIndex;
	};
	apply({ kind: 'angle', value: 54 });
	apply({ kind: 'stop', index: 0, position: 23, transparency: 37 });
	const inserted = apply({ kind: 'add', index: 0 });
	expect(inserted).toBe(2);
	apply({
		kind: 'stop',
		index: inserted,
		position: 56,
		color: { rgb: '00FF00' },
		transparency: 13,
	});
	const fill = chartOf(book).series[0]!.fill!;
	if (fill.kind !== 'gradient') throw new Error('Expected gradient');
	expect(fill.stops.map((stop) => stop.position)).toEqual([23, 100, 56]);
	expect(fill.sourceXml).toContain('rotWithShape="1"');
	apply({ kind: 'remove', index: 1 });
	expect(native.minimumRejected).toBe(true);
	expect(chartSeriesGradientPatch(chartOf(book), 0, { kind: 'remove', index: 0 })).toBeUndefined();
	const saved = chartOf(await loadXlsx(await saveXlsx(book)));
	const nativeSaved = chartOf(await setup(native.cases[3]!));
	if (
		saved.series[0]!.fill?.kind !== 'gradient' ||
		nativeSaved.series[0]!.fill?.kind !== 'gradient'
	)
		throw new Error('Expected gradients');
	expect(saved.series[0]!.fill.stops).toEqual(nativeSaved.series[0]!.fill.stops);
	session.undo();
	expect(chartOf(book).series[0]!.fill).toEqual(fill);
	for (const edit of [
		{ kind: 'angle', value: 361 },
		{ kind: 'stop', index: 0, position: -1 },
		{ kind: 'stop', index: 0, transparency: 101 },
		{ kind: 'stop', index: 99 },
	] as const)
		expect(() => chartSeriesGradientPatch(chartOf(book), 0, edit)).toThrow(RangeError);
});
