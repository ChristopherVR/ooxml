import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../../../chart/index';
import { setChartSeriesChartType } from '../../core/builders/sdk/chart-operations';
import { withChartSeriesOnSecondaryAxis } from '../../core/builders/sdk/chart-series-axis-operations';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../index';
import type { ChartPptxElement, PptxData } from '../../index';
import { requireFixture } from '../require-fixture';

/**
 * PowerPoint "Chart Filters" (limitations.md "Office chart extensions
 * (c15:/c16:/c16r3:)" row): confirms a series hidden via `Series.IsFiltered`
 * (COM ground truth, `chart-filtered-series.pptx`) is (1) modelled into
 * {@link PptxChartData.filteredSeries} rather than silently invisible, (2)
 * never counted among the plotted series, (3) preserved byte-identical when
 * the chart is untouched, and (4) does not collide with a newly added
 * visible series' `c:idx` after an edit (see chart-filtered-series.ts).
 */
const FIXTURE = requireFixture(
	fileURLToPath(new URL('../fixtures/e2e/chart-filtered-series.pptx', import.meta.url)),
);

function loadFixtureBytes(): Uint8Array {
	const buf = readFileSync(FIXTURE);
	return new Uint8Array(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

async function readZipEntry(bytes: Uint8Array, path: string): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	const file = zip.file(path);
	if (!file) {
		throw new Error(`missing ${path} in package`);
	}
	return file.async('string');
}

function chartOnSlide(data: PptxData, slideIndex: number): ChartPptxElement {
	const chart = data.slides[slideIndex]?.elements.find(
		(element): element is ChartPptxElement => element.type === 'chart',
	);
	if (!chart) {
		throw new Error(`no chart element on slide ${slideIndex + 1}`);
	}
	return chart;
}

describe('chart filtered-series round-trip (c15:filteredBarSeries)', () => {
	it('parses the hidden series into filteredSeries without plotting it', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(toArrayBuffer(loadFixtureBytes()));
		const chart = chartOnSlide(data, 0);

		// Series A and Series C are visible; Series B was filtered.
		expect(chart.chartData?.series.map((s) => s.name)).toStrictEqual(['Series A', 'Series C']);

		const filtered = chart.chartData?.filteredSeries;
		expect(filtered).toHaveLength(1);
		expect(filtered?.[0]).toMatchObject({
			idx: 1,
			order: 1,
			name: 'Series B',
			categories: ['Cat1', 'Cat2', 'Cat4'],
			values: [20, 21, 23],
		});
		expect(filtered?.[0]?.uniqueId).toMatch(/^\{[0-9A-F-]{36}\}$/u);
	});

	it('leaves an untouched chart part byte-identical, including the filter extension', async () => {
		const originalBytes = loadFixtureBytes();
		const handler = new PptxHandler();
		const data = await handler.load(toArrayBuffer(originalBytes));

		const savedBytes = await handler.save(data.slides);

		const originalXml = await readZipEntry(originalBytes, 'ppt/charts/chart1.xml');
		const savedXml = await readZipEntry(savedBytes, 'ppt/charts/chart1.xml');
		expect(savedXml).toBe(originalXml);
		expect(savedXml).toContain('c15:filteredBarSeries');
	});

	it('gives a new series added after the edit a c:idx that does not collide with the filtered series', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(toArrayBuffer(loadFixtureBytes()));
		const chart = chartOnSlide(data, 0);

		// Two visible series (idx 0, 2) plus the filtered one (idx 1). Add a
		// fourth (visible) series: before the fix, this would have reassigned
		// idx 0/1 to the two existing visible series (positional 0..1), landing
		// idx 1 on top of the still-untouched filtered series.
		chart.chartData!.series.push({ name: 'Series D', values: [40, 41, 42] });

		const savedBytes = await handler.save(data.slides);
		const savedXml = await readZipEntry(savedBytes, 'ppt/charts/chart1.xml');

		// The filtered series' own idx must still be exactly 1.
		const filteredSeriesXml = savedXml.slice(savedXml.indexOf('c15:filteredBarSeries'));
		expect(filteredSeriesXml).toContain('<c:idx val="1"');

		// Reload and confirm every VISIBLE c:ser has a distinct idx, none of
		// them 1 (reserved by the filtered series).
		const reloaded = await new PptxHandler().load(toArrayBuffer(savedBytes));
		const reloadedChart = chartOnSlide(reloaded, 0);
		expect(reloadedChart.chartData?.series.map((s) => s.name)).toStrictEqual([
			'Series A',
			'Series C',
			'Series D',
		]);
		// The filtered series must still round-trip untouched.
		expect(reloadedChart.chartData?.filteredSeries).toHaveLength(1);
		expect(reloadedChart.chartData?.filteredSeries?.[0]?.idx).toBe(1);

		const idxMatches = [...savedXml.matchAll(/<c:idx val="(\d+)"/gu)].map((m) => Number(m[1]));
		// Every c:idx value in the whole part (visible series + filtered
		// series) must be unique.
		expect(new Set(idxMatches).size).toBe(idxMatches.length);
	});

	// The chart editor's hide/show ("Chart Filters") actions
	// (pptx-viewer-shared's hideChartSeries/restoreFilteredSeries) actually
	// change PptxChartData.filteredSeries, so this now needs a writer
	// (chart-filtered-series-writer.ts) or the edit would silently vanish on
	// save: filteredSeries was previously read-mostly, round-tripping only as
	// opaque passthrough on an otherwise-untouched chart.
	describe('editing filteredSeries (hide/restore) survives save + reload', () => {
		it('restoring the filtered series into a real <c:ser> removes it from the c15 extension', async () => {
			const handler = new PptxHandler();
			const data = await handler.load(toArrayBuffer(loadFixtureBytes()));
			const chart = chartOnSlide(data, 0);
			const filtered = chart.chartData!.filteredSeries![0]!;

			// Mirrors pptx-viewer-shared's restoreFilteredSeries.
			chart.chartData!.series = [
				...chart.chartData!.series,
				{ name: filtered.name!, values: filtered.values! },
			];
			chart.chartData!.filteredSeries = undefined;

			const savedBytes = await handler.save(data.slides);
			const savedXml = await readZipEntry(savedBytes, 'ppt/charts/chart1.xml');
			expect(savedXml).not.toContain('c15:filteredBarSeries');

			const reloaded = await new PptxHandler().load(toArrayBuffer(savedBytes));
			const reloadedChart = chartOnSlide(reloaded, 0);
			expect(reloadedChart.chartData?.series.map((s) => s.name)).toStrictEqual([
				'Series A',
				'Series C',
				'Series B',
			]);
			expect(reloadedChart.chartData?.filteredSeries).toBeUndefined();
		});

		it('hiding a visible series writes it into a new c15:filteredBarSeries entry', async () => {
			const handler = new PptxHandler();
			const data = await handler.load(toArrayBuffer(loadFixtureBytes()));
			const chart = chartOnSlide(data, 0);
			const hiddenSeries = chart.chartData!.series[0]!; // "Series A"

			// Mirrors pptx-viewer-shared's hideChartSeries.
			chart.chartData!.series = chart.chartData!.series.filter((s) => s !== hiddenSeries);
			chart.chartData!.filteredSeries = [
				...chart.chartData!.filteredSeries!,
				{ idx: 2, order: 2, name: hiddenSeries.name, values: hiddenSeries.values },
			];

			const savedBytes = await handler.save(data.slides);
			const reloaded = await new PptxHandler().load(toArrayBuffer(await savedBytes));
			const reloadedChart = chartOnSlide(reloaded, 0);

			expect(reloadedChart.chartData?.series.map((s) => s.name)).toStrictEqual(['Series C']);
			const stillFiltered = reloadedChart.chartData?.filteredSeries ?? [];
			expect(stillFiltered.map((f) => f.name)).toStrictEqual(['Series B', 'Series A']);
			const restoredA = stillFiltered.find((f) => f.name === 'Series A');
			expect(restoredA?.values).toStrictEqual(hiddenSeries.values);
		});
	});
});

it('keeps a fully filtered line group and its workbook references in a combo chart', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank')
		.addChart('bar', {
			categories: ['Q1', 'Q2'],
			series: [
				{ name: 'Revenue', values: [10, 20] },
				{ name: 'Margin', values: [3, 5] },
			],
			embedWorkbook: true,
		})
		.build();
	data.slides.push(slide);
	const authored = chartOnSlide(data, 0);
	setChartSeriesChartType(authored, 1, 'line');
	authored.chartData = withChartSeriesOnSecondaryAxis(authored.chartData!, 1, true);
	const initial = await handler.save(data.slides);

	const editor = new PptxHandler();
	const opened = await editor.load(initial);
	const chart = chartOnSlide(opened, 0).chartData!;
	const [bar, line] = chart.series;
	expect(line?.seriesChartType).toBe('line');
	const part = chart.chartPartPath!;
	const before = parseChartSpace(await readZipEntry(initial, part)).chartSpace;
	const lineGroup = before.plotArea.groups[1]!;
	const range = lineGroup.series[0]!.values?.formula;
	expect(range).toBe('Sheet1!$C$2:$C$3');

	// Mirrors the shared editor hiding its only line series.
	chart.series = [bar!];
	chart.filteredSeries = [
		{
			idx: line!.idx!,
			order: line!.idx!,
			name: line!.name,
			categories: chart.categories,
			values: line!.values,
			originalSeries: line!,
		},
	];
	const hidden = await editor.save(opened.slides);
	const hiddenXml = await readZipEntry(hidden, part);
	const hiddenGroup = parseChartSpace(hiddenXml).chartSpace.plotArea.groups[1];
	expect(hiddenGroup).toMatchObject({ kind: 'line', axisIds: lineGroup.axisIds, series: [] });
	expect(hiddenGroup?.extLst).toContain('<c15:filteredLineSeries>');
	expect(hiddenGroup?.extLst).toContain(`<c15:sqref>${range}</c15:sqref>`);

	const reopen = new PptxHandler();
	const reopened = await reopen.load(hidden);
	const roundTripped = chartOnSlide(reopened, 0).chartData!;
	expect(roundTripped.series.map((series) => series.name)).toEqual(['Revenue']);
	const [entry] = roundTripped.filteredSeries ?? [];
	expect(entry).toMatchObject({ name: 'Margin', seriesChartType: 'line', axisId: line!.axisId });

	roundTripped.series.push({
		idx: entry!.idx,
		name: entry!.name!,
		values: entry!.values!,
		seriesChartType: entry!.seriesChartType,
		axisId: entry!.axisId,
	});
	roundTripped.filteredSeries = undefined;
	const restored = await reopen.save(reopened.slides);
	const restoredGroup = parseChartSpace(await readZipEntry(restored, part)).chartSpace.plotArea
		.groups[1];
	expect(restoredGroup).toMatchObject({
		kind: 'line',
		axisIds: lineGroup.axisIds,
		series: [{ index: line!.idx, values: { formula: range } }],
	});
	expect(restoredGroup?.extLst).toBeUndefined();

	const workbookPart = Object.keys((await JSZip.loadAsync(initial)).files).find((path) =>
		path.endsWith('.xlsx'),
	)!;
	const workbook = async (bytes: Uint8Array) =>
		(await JSZip.loadAsync(bytes)).file(workbookPart)!.async('uint8array');
	const initialWorkbook = await workbook(initial);
	expect(await workbook(hidden)).toEqual(initialWorkbook);
	expect(await workbook(restored)).toEqual(initialWorkbook);
});
