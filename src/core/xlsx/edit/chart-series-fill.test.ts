import JSZip from 'jszip';
import { expect, it } from 'vitest';
import native from '../__fixtures__/excel-chart-gradients.json';
import nativeFill from '../__fixtures__/excel-chart-series-fill.json';
import { NS, parseXml } from '../../xml/index';
import { chartView } from '../layout/chart-view';
import { THEME_SLOTS } from '../layout/colors';
import type { ChartObject } from '../model';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { chartDrawingColor, chartSeriesFillPatch } from './chart-series-fill';

async function setup() {
	const book = createWorkbook();
	book.theme.colors = THEME_SLOTS.map((slot) =>
		(native.scheme as Record<string, string>)[slot]!.slice(1),
	);
	createEditSession(book).addChart(0, {
		chartType: 'column',
		series: [],
		showLegend: true,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
	});
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	for (const [part, xml] of Object.entries(native.cases[0]!.parts)) zip.file(part, xml);
	return loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
}
const chartOf = (book: ReturnType<typeof createWorkbook>) =>
	book.sheets[0]!.drawings[0] as ChartObject;

it('matches Excel COM series fill brightness and replacement of point overrides', () => {
	const rgb = (value: number) =>
		`#${[value & 255, (value >> 8) & 255, (value >> 16) & 255]
			.map((channel) => channel.toString(16).padStart(2, '0'))
			.join('')
			.toUpperCase()}`;
	const book = createWorkbook();
	book.theme.colors[4] = rgb(nativeFill.accent1Rgb).slice(1);
	const chart: ChartObject = {
		kind: 'chart',
		chartType: 'column',
		showLegend: false,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		series: [
			{
				values: [1],
				categories: ['A'],
				pointColors: { 0: { kind: 'srgb', value: '00FF00', transforms: [] } },
			},
		],
	};
	const paint = chartSeriesFillPatch(chart, 0, { theme: 4, tint: 0.4 })!;
	const view = chartView(book, 0, { ...chart, ...paint }, () => []);
	expect(view.series[0]!.color).toBe(rgb(nativeFill.series[1]!.rgb));
	expect(nativeFill.series[1]!.brightness).toBeCloseTo(0.4, 6);
	expect(nativeFill.series[0]!.firstPointRgb).toBe(nativeFill.series[0]!.rgb);
	expect(paint.series![0]!.pointColors).toBeUndefined();
	const clear = chartSeriesFillPatch(chart, 0, null)!;
	expect(clear.series![0]!.pointColors).toBeUndefined();
	expect(nativeFill.series[2]!.firstPointVisible).toBe(nativeFill.series[2]!.visible);
	expect(nativeFill.series[2]!.visible).toBe(0);
});

it('edits one native gradient series, preserving other series, axes and effects through undo and save', async () => {
	const book = await setup();
	const session = createEditSession(book);
	const before = structuredClone(chartOf(book));
	const patch = chartSeriesFillPatch(chartOf(book), 1, { theme: 4, tint: 0.4 })!;
	session.updateChart(0, 0, patch);
	const edited = chartOf(book);
	expect(edited.series[0]).toEqual(before.series[0]);
	expect(edited.series[1]!.fill).toBeUndefined();
	expect(edited.series[1]!.effectsXml).toBe(before.series[1]!.effectsXml);
	expect(edited.series[1]!.drawingColor).toEqual({
		kind: 'scheme',
		value: 'accent1',
		transforms: [
			{ name: 'lumMod', value: '60000' },
			{ name: 'lumOff', value: '40000' },
		],
	});
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	const xml = parseXml(await zip.file(edited.partName!)!.async('string'));
	const original = parseXml(new TextDecoder().decode(book.source!.parts.get(edited.partName!)!));
	for (const tag of ['catAx', 'valAx'])
		expect(Array.from(xml.getElementsByTagNameNS(NS.c, tag), (node) => node.outerHTML)).toEqual(
			Array.from(original.getElementsByTagNameNS(NS.c, tag), (node) => node.outerHTML),
		);
	expect(xml.getElementsByTagNameNS(NS.c, 'ser')[0]!.outerHTML).toBe(
		original.getElementsByTagNameNS(NS.c, 'ser')[0]!.outerHTML,
	);
	const restored = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(chartOf(restored).series[1]!.drawingColor).toEqual(edited.series[1]!.drawingColor);
	session.undo();
	expect(chartOf(book)).toEqual(before);
	session.redo();
	expect(chartOf(book).series[1]!.drawingColor).toEqual(edited.series[1]!.drawingColor);
});

it('clears point paints like Excel series fill edits, retains effects, and round-trips no fill', async () => {
	const book = await setup();
	const session = createEditSession(book);
	let chart = chartOf(book);
	const points = structuredClone(chart.series);
	points[1]!.pointColors = { 0: { kind: 'srgb', value: '00FF00', transforms: [] } };
	session.updateChart(0, 0, { series: points });
	chart = chartOf(book);
	session.updateChart(0, 0, chartSeriesFillPatch(chart, 1, null)!);
	expect(chart.series[1]!.pointColors).toBeUndefined();
	expect(chartView(book, 0, chart, () => []).series[1]!.color).toBe('none');
	const saved = await loadXlsx(await saveXlsx(book));
	expect(chartOf(saved).series[1]!.fill).toEqual({ kind: 'none' });
	expect(chartOf(saved).series[1]!.effectsXml).toBe(chart.series[1]!.effectsXml);
	expect(chartSeriesFillPatch(chart, 1, null)).toBeUndefined();
});

it('writes authored legacy theme tints without dropping their luminance transforms', async () => {
	const book = createWorkbook();
	createEditSession(book).addChart(0, {
		chartType: 'column',
		series: [{ values: [1], categories: ['A'], color: { theme: 4, tint: -0.25 } }],
		showLegend: false,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
	});
	const restored = await loadXlsx(await saveXlsx(book));
	expect(chartOf(restored).series[0]!.drawingColor).toEqual(
		chartDrawingColor({ theme: 4, tint: -0.25 }),
	);
	for (const color of [{ theme: 99 }, { rgb: 'garbage' }, { theme: 4, tint: NaN }])
		expect(chartDrawingColor(color)).toBeUndefined();
	expect(() => chartSeriesFillPatch(chartOf(book), -1, null)).toThrow(RangeError);
});
