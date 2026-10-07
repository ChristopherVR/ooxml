import JSZip from 'jszip';
import { expect, it } from 'vitest';
import native from '../__fixtures__/excel-chart-spacing.json';
import { NS, parseXml } from '../../xml/index';
import { createWorkbook } from '../workbook';
import { parseChart } from '../read/chart';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import type { ChartObject } from '../model';
import { chartView } from './chart-view';
import { renderChartSvg } from './chart-svg';
import { THEME_SLOTS } from './colors';

const sampleChart = (sample: (typeof native.cases)[number]) =>
	parseChart(
		sample.parts['xl/charts/chart1.xml'],
		{ from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		'xl/charts/chart1.xml',
	);
const theme = (book: ReturnType<typeof createWorkbook>) => {
	book.theme.colors = THEME_SLOTS.map((slot) =>
		(native.scheme as Record<string, string>)[slot]!.slice(1),
	);
};
const chartOf = (book: ReturnType<typeof createWorkbook>) =>
	book.sheets[0]!.drawings[0] as ChartObject;

for (const sample of native.cases)
	it(`matches COM bar geometry ${sample.type}/${sample.gapWidth}/${sample.overlap}`, () => {
		const book = createWorkbook();
		theme(book);
		const chart = sampleChart(sample);
		expect(chart.barGapWidth).toBe(sample.gapWidth);
		expect(chart.barOverlap).toBe(sample.overlap);
		const svg = parseXml(
			renderChartSvg(
				chartView(book, 0, chart, () => []),
				640,
				400,
			),
		);
		const groups = [...svg.getElementsByTagName('g')];
		const rect = (series: number, point: number) =>
			groups
				.find(
					(group) =>
						group.getAttribute('data-chart-series') === String(series) &&
						group.getAttribute('data-chart-point') === String(point),
				)!
				.getElementsByTagName('rect')[0]!;
		const position = sample.type === 57 ? 'y' : 'x';
		const dimension = sample.type === 57 ? 'height' : 'width';
		const nativePosition = sample.type === 57 ? 'top' : 'left';
		const nativeDimension = sample.type === 57 ? 'height' : 'width';
		const first = rect(0, 0),
			second = rect(0, 1),
			other = rect(1, 0);
		const pitch = Number(second.getAttribute(position)) - Number(first.getAttribute(position));
		const p0 = sample.points.find((p) => p.series === 0 && p.point === 0)!;
		const p1 = sample.points.find((p) => p.series === 0 && p.point === 1)!;
		const pOther = sample.points.find((p) => p.series === 1 && p.point === 0)!;
		const nativePitch = p1[nativePosition] - p0[nativePosition];
		expect(Math.sign(pitch)).toBe(Math.sign(nativePitch));
		expect(Number(first.getAttribute(dimension)) / Math.abs(pitch)).toBeCloseTo(
			p0[nativeDimension] / Math.abs(nativePitch),
			3,
		);
		expect(
			(Number(other.getAttribute(position)) - Number(first.getAttribute(position))) / pitch,
		).toBeCloseTo((pOther[nativePosition] - p0[nativePosition]) / nativePitch, 3);
	});

async function imported() {
	const book = createWorkbook();
	theme(book);
	createEditSession(book).addChart(0, { ...sampleChart(native.cases[0]!), series: [] });
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	for (const [path, xml] of Object.entries(native.cases[0]!.parts)) zip.file(path, xml);
	return loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
}

it('edits spacing in place, preserves source detail and supports regeneration and undo', async () => {
	const book = await imported();
	const session = createEditSession(book);
	const before = structuredClone(chartOf(book));
	session.updateChart(0, 0, { barGapWidth: 5, barOverlap: 23 });
	const bytes = await saveXlsx(book);
	const after = await loadXlsx(bytes);
	expect(chartOf(after).barGapWidth).toBe(5);
	expect(chartOf(after).barOverlap).toBe(23);
	const xml = parseXml(
		await (await JSZip.loadAsync(bytes)).file('xl/charts/chart1.xml')!.async('string'),
	);
	const original = parseXml(native.cases[0]!.parts['xl/charts/chart1.xml']);
	for (const element of ['ser', 'valAx', 'catAx'])
		expect([...xml.getElementsByTagNameNS(NS.c, element)].map(String)).toEqual(
			[...original.getElementsByTagNameNS(NS.c, element)].map(String),
		);
	session.undo();
	expect(chartOf(book)).toEqual(before);
	session.redo();
	session.updateChart(0, 0, { chartType: 'bar' });
	const regenerated = await loadXlsx(await saveXlsx(book));
	expect(chartOf(regenerated).barGapWidth).toBe(5);
	expect(chartOf(regenerated).barOverlap).toBe(23);
});

it('rejects out-of-range or noninteger spacing before changing history', async () => {
	const book = await imported();
	const original = structuredClone(chartOf(book));
	const session = createEditSession(book);
	for (const patch of [
		{ barGapWidth: 501 },
		{ barGapWidth: NaN },
		{ barOverlap: -101 },
		{ barOverlap: 0.5 },
	])
		expect(() => session.updateChart(0, 0, patch)).toThrow(RangeError);
	expect(chartOf(book)).toEqual(original);
});
