import JSZip from 'jszip';
import { expect, it } from 'vitest';
import native from '../__fixtures__/excel-chart-gradients.json';
import { NS, parseXml } from '../../xml/index';
import { chartView } from '../layout/chart-view';
import { renderChartSvg } from '../../chart/render/chart-svg';
import { THEME_SLOTS } from '../layout/colors';
import type { ChartObject } from '../model';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createWorkbook } from '../workbook';
import { chartColorPalettePatch } from './chart-colors';
import { createEditSession } from './session';

async function workbookFrom(sample = native.cases[0]!) {
	const book = createWorkbook();
	book.theme.colors = THEME_SLOTS.map((slot) =>
		(native.scheme as Record<string, string>)[slot]!.slice(1),
	);
	createEditSession(book).addChart(0, {
		chartType: 'column',
		showLegend: true,
		series: [],
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
	});
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	for (const [name, xml] of Object.entries(sample.parts)) zip.file(name, xml);
	return loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
}
const chartOf = (book: ReturnType<typeof createWorkbook>) =>
	book.sheets[0]!.drawings[0] as ChartObject;
const paints = (book: ReturnType<typeof createWorkbook>) =>
	chartView(book, 0, chartOf(book), () => []).series.map((series) =>
		series.gradient?.stops.map((stop) => ({
			color: stop.color.toUpperCase(),
			position: stop.position / 100,
			transparency: 1 - (stop.opacity ?? 1),
		})),
	);
const nativePaints = (sample: (typeof native.cases)[number]) =>
	sample.series.map((series) => series.stops);

for (const sample of native.cases) {
	it(`renders native palette ${sample.palette} gradient stops measured by Excel COM`, async () => {
		const book = await workbookFrom(sample);
		expect(paints(book)).toEqual(nativePaints(sample));
		const view = chartView(book, 0, chartOf(book), () => []);
		const xml = parseXml(renderChartSvg(view, 640, 400));
		const gradients = [...xml.getElementsByTagName('linearGradient')];
		expect(gradients).toHaveLength(2);
		for (const gradient of gradients)
			expect(
				[...xml.getElementsByTagName('rect')].some(
					(rect) => rect.getAttribute('fill') === `url(#${gradient.getAttribute('id')})`,
				),
			).toBe(true);
		expect(view.series[0]!.color).not.toContain('url(');
	});
	it(`changes palette 10 to ${sample.palette} without losing native gradient flags or effects`, async () => {
		const book = await workbookFrom();
		const before = structuredClone(chartOf(book));
		const session = createEditSession(book);
		session.updateChart(0, 0, chartColorPalettePatch(chartOf(book), sample.palette)!);
		expect(paints(book)).toEqual(nativePaints(sample));
		const bytes = await saveXlsx(book);
		const saved = await loadXlsx(bytes);
		expect(paints(saved)).toEqual(nativePaints(sample));
		const zip = await JSZip.loadAsync(bytes);
		const xml = parseXml(await zip.file('xl/charts/chart1.xml')!.async('string'));
		const original = parseXml(native.cases[0]!.parts['xl/charts/chart1.xml']);
		for (const name of ['outerShdw', 'lin', 'valAx'])
			expect(
				[...xml.getElementsByTagNameNS(name === 'valAx' ? NS.c : NS.a, name)].map(String),
			).toEqual(
				[...original.getElementsByTagNameNS(name === 'valAx' ? NS.c : NS.a, name)].map(String),
			);
		session.undo();
		expect(chartOf(book)).toEqual(before);
		session.redo();
		expect(paints(book)).toEqual(nativePaints(sample));
	});
}

it('writes, replaces and removes point fills while retaining native series effects', async () => {
	const book = await workbookFrom();
	const session = createEditSession(book);
	let series = structuredClone(chartOf(book).series);
	series[0]!.pointFills = { 1: structuredClone(series[0]!.fill!) };
	session.updateChart(0, 0, { series });
	let saved = await loadXlsx(await saveXlsx(book));
	expect(chartOf(saved).series[0]!.pointFills?.[1]?.kind).toBe('gradient');
	series = structuredClone(chartOf(book).series);
	series[0]!.pointColors = { 1: { kind: 'srgb', value: '123456', transforms: [] } };
	session.updateChart(0, 0, { series });
	saved = await loadXlsx(await saveXlsx(book));
	expect(chartOf(saved).series[0]!.pointFills?.[1]).toBeUndefined();
	expect(chartOf(saved).series[0]!.pointColors?.[1]?.value).toBe('123456');
	series = structuredClone(chartOf(book).series);
	delete series[0]!.pointColors;
	session.updateChart(0, 0, { series });
	saved = await loadXlsx(await saveXlsx(book));
	expect(chartOf(saved).series[0]!.pointColors?.[1]).toBeUndefined();
	expect(chartOf(saved).series[0]!.fill?.kind).toBe('gradient');
});

it('replaces gradients with explicit RGB edits and preserves gradients on chart regeneration', async () => {
	const book = await workbookFrom();
	const session = createEditSession(book);
	session.updateChart(0, 0, { chartType: 'bar' });
	let saved = await loadXlsx(await saveXlsx(book));
	expect(paints(saved)).toEqual(nativePaints(native.cases[0]!));
	expect(chartOf(saved).series.map((series) => series.effectsXml)).toEqual(
		chartOf(book).series.map((series) => series.effectsXml),
	);
	const series = structuredClone(chartOf(book).series);
	series[0]!.color = { rgb: '123456' };
	session.updateChart(0, 0, { series });
	saved = await loadXlsx(await saveXlsx(book));
	expect(chartOf(saved).series[0]!.fill).toBeUndefined();
	expect(chartView(saved, 0, chartOf(saved), () => []).series[0]!.color).toBe('#123456');
});

it('removes imported series effects without discarding the fill or other chart detail', async () => {
	const book = await workbookFrom();
	const session = createEditSession(book);
	const original = structuredClone(chartOf(book));
	const series = structuredClone(original.series);
	delete series[0]!.effectsXml;
	session.updateChart(0, 0, { series });
	const saved = await loadXlsx(await saveXlsx(book));
	expect(chartOf(saved).series[0]!.effectsXml).toBeUndefined();
	expect(chartOf(saved).series[1]!.effectsXml).toBe(original.series[1]!.effectsXml);
	expect(paints(saved)).toEqual(nativePaints(native.cases[0]!));
	session.undo();
	expect(chartOf(book)).toEqual(original);
});
