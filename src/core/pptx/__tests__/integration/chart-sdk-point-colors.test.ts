/**
 * `ChartSeriesInput.pointColors` on new charts, saved and reloaded: a filled point (bar, the bar
 * series of a combination chart, a doughnut slice) is written as `c:dPt/c:spPr`, beside a series
 * colour, and a point of a series drawn with markers as its marker's fill (`c:dPt/c:marker`);
 * `'none'` is written and reloaded as no fill.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { createChartElement } from '../../core/builders/sdk/ElementFactory';
import type { ChartInput } from '../../core/builders/sdk/types';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement, PptxChartType } from '../../core/types';
import { writeChartShapeProps } from '../../core/utils/chart-shape-props-writer';
import { firstElement, newChartDeck } from '../helpers/new-chart-deck';

const CATEGORIES = ['a', 'b', 'c'];

async function saved(type: PptxChartType, input: Partial<ChartInput>) {
	const { save } = await newChartDeck(type, { categories: CATEGORIES, ...input });
	const result = await save();
	expect(result.warnings).toStrictEqual([]);
	return result;
}

const seriesXml = (xml: string): string[] => xml.match(/<c:ser>.*?<\/c:ser>/gsu) ?? [];
const pointsXml = (xml: string): string[] => xml.match(/<c:dPt>.*?<\/c:dPt>/gsu) ?? [];

describe('point colours of new charts', () => {
	it('writes no point override when pointColors is omitted or empty', async () => {
		for (const pointColors of [undefined, []]) {
			const { xml, chart } = await saved('bar', {
				series: [
					{ name: 'A', values: [1, 2, 3], color: '#3B5BDB', ...(pointColors && { pointColors }) },
				],
			});
			expect(xml).not.toContain('<c:dPt>');
			expect(chart.series[0]!.dataPoints ?? []).toStrictEqual([]);
		}
	});

	it('keeps a bar point colour beside the series colour', async () => {
		const { xml, chart } = await saved('bar', {
			series: [
				{ name: 'A', values: [1, 2, 3], color: '#3B5BDB', pointColors: [null, null, '#DBE4FF'] },
			],
		});
		const series = seriesXml(xml)[0]!;
		expect(firstElement(series, 'c:spPr')).toContain('<a:srgbClr val="3B5BDB">');
		const points = pointsXml(series);
		expect(points).toHaveLength(1);
		expect(points[0]).toMatch(
			/^<c:dPt><c:idx val="2"><\/c:idx>(<c:invertIfNegative[^>]*><\/c:invertIfNegative>)?<c:spPr><a:solidFill><a:srgbClr val="DBE4FF">/u,
		);
		expect(chart.series[0]!.color).toBe('#3B5BDB');
		expect(chart.series[0]!.dataPoints).toStrictEqual([
			expect.objectContaining({ idx: 2, spPr: { fillColor: '#DBE4FF' } }),
		]);
	});

	it('keeps a bar point colour beside a palette series colour', async () => {
		const { xml, chart } = await saved('bar', {
			palette: ['#112233'],
			series: [{ name: 'A', values: [1, 2, 3], pointColors: ['#DBE4FF'] }],
		});
		expect(firstElement(seriesXml(xml)[0]!, 'c:spPr')).toContain('<a:srgbClr val="112233">');
		expect(pointsXml(xml)[0]).toContain('<a:srgbClr val="DBE4FF">');
		expect(chart.series[0]!.dataPoints).toStrictEqual([
			expect.objectContaining({ idx: 0, spPr: { fillColor: '#DBE4FF' } }),
		]);
	});

	it('fills the marker of a line point, keeping the series marker', async () => {
		const marker = {
			symbol: 'circle' as const,
			size: 7,
			fill: '#FFFFFF',
			line: { color: '#3B5BDB', width: 1.5 },
		};
		const { xml, chart } = await saved('line', {
			series: [
				{ name: 'L', values: [1, 2, 3], color: '#3B5BDB', marker, pointColors: ['#E03131'] },
			],
		});
		const points = pointsXml(seriesXml(xml)[0]!);
		expect(points).toHaveLength(1);
		const point = points[0]!;
		expect(point).toContain('<c:idx val="0">');
		expect(firstElement(point, 'c:marker')).toMatch(
			/<c:symbol val="circle"><\/c:symbol><c:size val="7"><\/c:size><c:spPr><a:solidFill><a:srgbClr val="E03131"><\/a:srgbClr><\/a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="3B5BDB">/u,
		);
		// PowerPoint draws no fill of a line point's own c:spPr, so none is written.
		expect(point.replace(firstElement(point, 'c:marker'), '')).not.toContain('E03131');
		expect(chart.series[0]!.dataPoints).toStrictEqual([
			{
				idx: 0,
				marker: {
					symbol: 'circle',
					size: 7,
					spPr: { fillColor: '#E03131', strokeColor: '#3B5BDB', strokeWidth: 1.5 },
				},
			},
		]);
	});

	it('writes an automatic symbol for a line point when the series has no marker', async () => {
		const { xml, chart } = await saved('line', {
			series: [{ name: 'L', values: [1, 2, 3], pointColors: [null, '#E03131'] }],
		});
		expect(firstElement(pointsXml(xml)[0]!, 'c:marker')).toContain('<c:symbol val="auto">');
		expect(chart.series[0]!.dataPoints).toStrictEqual([
			{ idx: 1, marker: { symbol: 'auto', spPr: { fillColor: '#E03131' } } },
		]);
	});

	it('paints each series of a combination chart as its own type draws it', async () => {
		const { xml, chart } = await saved('bar', {
			series: [
				{ name: 'A', values: [1, 2, 3], color: '#3B5BDB', pointColors: [null, null, '#DBE4FF'] },
				{
					name: 'B',
					values: [2, 2, 2],
					chartType: 'line',
					color: '#868E96',
					marker: { symbol: 'circle', size: 6 },
					pointColors: ['#E03131'],
				},
			],
		});
		expect(xml).toContain('<c:barChart>');
		expect(xml).toContain('<c:lineChart>');
		const [bar, line] = seriesXml(xml);
		expect(pointsXml(bar!)[0]).toContain('<c:spPr><a:solidFill><a:srgbClr val="DBE4FF">');
		expect(firstElement(pointsXml(line!)[0]!, 'c:marker')).toContain('<a:srgbClr val="E03131">');
		const [barModel, lineModel] = chart.series;
		expect(barModel!.dataPoints).toStrictEqual([
			expect.objectContaining({ idx: 2, spPr: { fillColor: '#DBE4FF' } }),
		]);
		expect(lineModel!.dataPoints).toStrictEqual([
			{ idx: 0, marker: { symbol: 'circle', size: 6, spPr: { fillColor: '#E03131' } } },
		]);
	});

	it('keeps doughnut slice colours beside a series colour and a palette', async () => {
		const { xml, chart } = await saved('doughnut', {
			palette: ['#112233'],
			series: [{ name: 'D', values: [1, 2, 3], color: '#3B5BDB', pointColors: [null, '#DBE4FF'] }],
		});
		const points = pointsXml(xml);
		expect(points.map((point) => /srgbClr val="(\w+)"/u.exec(point)?.[1])).toStrictEqual([
			'112233',
			'DBE4FF',
			'112233',
		]);
		expect(chart.series[0]!.dataPoints!.map((point) => point.spPr?.fillColor)).toStrictEqual([
			'#112233',
			'#DBE4FF',
			'#112233',
		]);
	});

	it('refuses more colours than values, a malformed colour and a series without markers', () => {
		const one = (series: ChartInput['series'][number]): ChartInput => ({
			categories: CATEGORIES,
			series: [series],
		});
		expect(() =>
			createChartElement(
				'bar',
				one({ name: 'A', values: [1, 2], pointColors: ['#111111', null, '#222222'] }),
			),
		).toThrow(RangeError);
		expect(() =>
			createChartElement('bar', one({ name: 'A', values: [1, 2, 3], pointColors: ['red'] })),
		).toThrow(/Point colour/u);
		expect(() =>
			createChartElement(
				'line',
				one({ name: 'L', values: [1, 2, 3], marker: { symbol: 'none' }, pointColors: ['#111111'] }),
			),
		).toThrow(/markers/u);
	});
});

describe("'none' point colours of new charts", () => {
	it('writes and reloads a bar point and a doughnut slice without fill', async () => {
		for (const type of ['bar', 'doughnut'] as const) {
			const { xml, chart } = await saved(type, {
				series: [
					{
						name: 'A',
						values: [1, 2, 3],
						color: '#3B5BDB',
						pointColors: ['none', null, '#DBE4FF'],
					},
				],
			});
			const points = pointsXml(xml);
			expect(points).toHaveLength(2);
			expect(firstElement(points[0]!, 'c:spPr')).toBe('<c:spPr><a:noFill></a:noFill></c:spPr>');
			expect(points[1]).toContain('<a:srgbClr val="DBE4FF">');
			expect(chart.series[0]!.dataPoints!.map((point) => point.spPr)).toStrictEqual([
				{ fillColor: 'none' },
				{ fillColor: '#DBE4FF' },
			]);
		}
	});

	it('writes and reloads a line point marker without fill', async () => {
		const { xml, chart } = await saved('line', {
			series: [
				{
					name: 'L',
					values: [1, 2, 3],
					marker: { symbol: 'circle', size: 7, line: { color: '#3B5BDB' } },
					pointColors: [null, 'none'],
				},
			],
		});
		expect(firstElement(pointsXml(xml)[0]!, 'c:marker')).toContain(
			'<c:spPr><a:noFill></a:noFill><a:ln><a:solidFill><a:srgbClr val="3B5BDB">',
		);
		expect(chart.series[0]!.dataPoints).toStrictEqual([
			{
				idx: 1,
				marker: { symbol: 'circle', size: 7, spPr: { fillColor: 'none', strokeColor: '#3B5BDB' } },
			},
		]);
	});

	it('keeps the reloaded none fill through a save that edits something else', async () => {
		const { bytes } = await saved('bar', {
			series: [{ name: 'A', values: [1, 2, 3], pointColors: ['none'] }],
		});
		const handler = new PptxHandler();
		const loaded = await handler.load(bytes.buffer as ArrayBuffer);
		const element = loaded.slides[0]!.elements.find(
			(item) => item.type === 'chart',
		) as ChartPptxElement;
		element.chartData!.title = 'Edited';
		const again = await handler.save(loaded.slides);
		const xml = await (await JSZip.loadAsync(again)).file('ppt/charts/chart1.xml')!.async('string');
		expect(xml).toContain('Edited');
		expect(pointsXml(xml)).toStrictEqual([
			expect.stringMatching(/<c:idx val="0"\/>.*<c:spPr><a:noFill\/><\/c:spPr>/su),
		]);
		expect(xml).not.toContain('val="none"');
	});

	it('writes a none point fill through the object-tree writer as a:noFill', () => {
		const spPr = writeChartShapeProps(
			{ 'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } } },
			{ fillColor: 'none' },
			(key) => key.replace(/^a:/u, ''),
		);
		expect(spPr).toStrictEqual({ 'a:noFill': {} });
	});
});
