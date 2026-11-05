/**
 * The declarative chart inputs of `addChart` (`ChartInput`, `ChartSeriesInput`, the axis edit):
 * each case builds a deck with the SDK, saves it and reads the chart part, so the fields are
 * proven to reach the file without post-processing.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { createChartElement } from '../../core/builders/sdk/ElementFactory';
import type { ChartInput } from '../../core/builders/sdk/types';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement, PptxChartData, PptxChartType } from '../../core/types';
import { writeNewChartPart } from '../../core/utils/chart-new-part';

const CATEGORIES = ['Q1', 'Q2', 'Q3'];

async function save(
	type: PptxChartType,
	input: ChartInput,
): Promise<{ xml: string; chart: PptxChartData }> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank')
		.addChart(type, input, { x: 50, y: 50, width: 500, height: 300 })
		.build();
	data.slides.push(slide);
	const element = slide.elements.find((item) => item.type === 'chart') as ChartPptxElement;
	// The shared chart core writes the part in full; the legacy generator would drop these fields.
	expect(writeNewChartPart(structuredClone(element.chartData!))).toMatchObject({ unwritten: [] });
	const saved = await handler.save(data.slides);
	const xml = await (await JSZip.loadAsync(saved)).file('ppt/charts/chart1.xml')!.async('string');
	const loaded = await new PptxHandler().load(saved.buffer as ArrayBuffer);
	const chart = loaded.slides[0]!.elements.find((item) => item.type === 'chart') as
		| ChartPptxElement
		| undefined;
	return { xml, chart: chart!.chartData! };
}

/** The first `c:<name>` element of `xml`, or of each match when `all`. */
const element = (xml: string, name: string): string =>
	new RegExp(`<c:${name}>.*?</c:${name}>`, 'su').exec(xml)?.[0] ?? '';
const elements = (xml: string, name: string): string[] =>
	xml.match(new RegExp(`<c:${name}>.*?</c:${name}>`, 'gsu')) ?? [];

describe('declarative chart inputs', () => {
	it('writes blank points and dispBlanksAs', async () => {
		const { xml, chart } = await save('line', {
			categories: CATEGORIES,
			dispBlanksAs: 'span',
			series: [{ name: 'A', values: [1, null, 3] }],
		});
		expect(xml).toContain('<c:dispBlanksAs val="span">');
		const values = element(xml, 'val');
		expect(values).toContain('<c:ptCount val="3">');
		expect(values).not.toContain('<c:pt idx="1">');
		expect(chart.series[0]!.blanks).toStrictEqual([false, true, false]);
	});

	it('writes a combination chart with a line series on a secondary axis', async () => {
		const { xml, chart } = await save('bar', {
			categories: CATEGORIES,
			series: [
				{ name: 'Revenue', values: [10, 20, 30] },
				{ name: 'Margin', values: [0.1, 0.2, 0.3], chartType: 'line', secondaryAxis: true },
			],
		});
		const bar = element(xml, 'barChart');
		const line = element(xml, 'lineChart');
		expect(bar).toContain('<c:v>Revenue</c:v>');
		expect(bar).not.toContain('Margin');
		expect(line).toContain('<c:v>Margin</c:v>');
		const valueAxes = elements(xml, 'valAx');
		expect(valueAxes).toHaveLength(2);
		const secondary = valueAxes.find((axis) => axis.includes('<c:axPos val="r">'))!;
		expect(secondary).toContain('<c:crosses val="max">');
		const secondaryId = /<c:axId val="(\d+)">/u.exec(secondary)![1];
		expect(line).toContain(`<c:axId val="${secondaryId}">`);
		expect(bar).toContain('<c:axId val="222222222">');
		expect(chart.chartType).toBe('combo');
		expect(chart.series.map((series) => series.seriesChartType)).toStrictEqual(['bar', 'line']);
	});

	it('writes series line width, dash, smoothing and a marker', async () => {
		const { xml } = await save('line', {
			categories: CATEGORIES,
			series: [
				{
					name: 'A',
					values: [1, 2, 3],
					color: '#C00000',
					line: { width: 2.5, dash: 'dash' },
					smooth: true,
					marker: { symbol: 'diamond', size: 8, fill: '#00B050', line: 'none' },
				},
			],
		});
		const series = element(xml, 'ser');
		expect(series).toContain('<a:ln w="31750">');
		expect(series).toContain('<a:prstDash val="dash">');
		expect(series).toContain('<c:smooth val="1">');
		const marker = element(series, 'marker');
		expect(marker).toContain('<c:symbol val="diamond">');
		expect(marker).toContain('<c:size val="8">');
		expect(marker).toContain('<a:srgbClr val="00B050">');
		expect(marker).toMatch(/<a:ln>\s*<a:noFill/u);
	});

	it('writes series data labels with number format, font, rotation, wrap and fill', async () => {
		const { xml, chart } = await save('bar', {
			categories: CATEGORIES,
			dataLabels: { showValue: true, font: { size: 9, color: '#404040' } },
			series: [
				{
					name: 'A',
					values: [0.1, 0.25, 0.5],
					dataLabels: {
						position: 'outEnd',
						numberFormat: '0%',
						font: { bold: true },
						rotation: -45,
						wrap: false,
						fill: '#FFFF00',
					},
				},
				{ name: 'B', values: [1, 2, 3] },
			],
		});
		const [first, second] = elements(xml, 'ser').map((series) => element(series, 'dLbls'));
		expect(first).toContain('<c:numFmt formatCode="0%" sourceLinked="0">');
		expect(first).toContain('<c:dLblPos val="outEnd">');
		expect(first).toContain('<c:showVal val="1">');
		expect(first).toMatch(/<a:bodyPr[^>]* rot="-2700000"/u);
		expect(first).toMatch(/<a:bodyPr[^>]* wrap="none"/u);
		expect(first).toMatch(/<a:defRPr[^>]* sz="900"[^>]* b="1"|<a:defRPr[^>]* b="1"[^>]* sz="900"/u);
		expect(first).toContain('<a:srgbClr val="FFFF00">');
		expect(second).toContain('<c:showVal val="1">');
		expect(second).toContain('sz="900"');
		expect(second).not.toContain('dLblPos');
		expect(chart.series[0]!.dataLabelOptions).toMatchObject({
			showValue: true,
			position: 'outEnd',
			numberFormat: '0%',
		});
		expect(
			Object.entries(chart.elementFormats ?? {}).filter(([key]) => key.startsWith('dataLabels')),
		).toStrictEqual([
			[
				expect.any(String),
				expect.objectContaining({
					text: expect.objectContaining({ rotation: -45, wrap: false }),
				}),
			],
		]);
	});

	it('writes scatter X values, bubble sizes and the scatter style', async () => {
		const scatter = await save('scatter', {
			categories: CATEGORIES,
			scatterStyle: 'marker',
			series: [{ name: 'S', values: [5, 6, 7], xValues: [1.5, 2.5, 4] }],
		});
		expect(scatter.xml).toContain('<c:scatterStyle val="marker">');
		expect(element(scatter.xml, 'xVal')).toMatch(
			/<c:pt idx="0"><c:v>1.5<\/c:v>.*<c:pt idx="2"><c:v>4<\/c:v>/u,
		);
		expect(scatter.chart.series[0]!.xValues).toStrictEqual([1.5, 2.5, 4]);
		const bubble = await save('bubble', {
			categories: CATEGORIES,
			series: [{ name: 'B', values: [5, 6, 7], xValues: [1, 2, 3], bubbleSizes: [10, 20, 30] }],
		});
		expect(element(bubble.xml, 'bubbleSize')).toMatch(
			/<c:numLit>.*<c:pt idx="1"><c:v>20<\/c:v>.*<c:pt idx="2"><c:v>30<\/c:v>/u,
		);
		expect(bubble.chart.series[0]!.bubbleSizes).toStrictEqual([10, 20, 30]);
	});

	it('expands a palette into point fills on a doughnut and series colours on a bar', async () => {
		const doughnut = await save('doughnut', {
			categories: CATEGORIES,
			palette: ['#112233', '#445566'],
			series: [{ name: 'D', values: [1, 2, 3], pointColors: [null, '#ABCDEF'] }],
		});
		const points = elements(doughnut.xml, 'dPt');
		expect(points).toHaveLength(3);
		expect(points[0]).toContain('<a:srgbClr val="112233">');
		expect(points[1]).toContain('<a:srgbClr val="ABCDEF">');
		expect(points[2]).toContain('<a:srgbClr val="112233">');
		const bar = await save('bar', {
			categories: CATEGORIES,
			palette: ['#112233', '#445566'],
			series: [
				{ name: 'A', values: [1, 2, 3] },
				{ name: 'B', values: [1, 2, 3] },
				{ name: 'C', values: [1, 2, 3], color: '#000000' },
			],
		});
		const colors = elements(bar.xml, 'ser').map(
			(series) => /<c:spPr><a:solidFill><a:srgbClr val="(\w+)"/u.exec(series)?.[1],
		);
		expect(colors).toStrictEqual(['112233', '445566', '000000']);
		expect(bar.xml).not.toContain('<c:dPt>');
	});

	it('writes a series with no fill', async () => {
		const { xml } = await save('bar', {
			categories: CATEGORIES,
			series: [{ name: 'A', values: [1, 2, 3], fill: 'none' }],
		});
		expect(element(element(xml, 'ser'), 'spPr')).toMatch(/^<c:spPr><a:noFill/u);
	});

	it('writes axis title, number format, font, rotation and line', async () => {
		const { xml } = await save('bar', {
			categories: CATEGORIES,
			series: [{ name: 'A', values: [1000, 2000, 3000] }],
			axes: {
				valAx: { title: 'Revenue', numberFormat: '#,##0', font: { size: 8, color: '#595959' } },
				catAx: { rotation: -45, line: 'none', font: { italic: true } },
			},
		});
		const value = element(xml, 'valAx');
		expect(value).toContain('<a:t>Revenue</a:t>');
		expect(value).toContain('<c:numFmt formatCode="#,##0" sourceLinked="0">');
		expect(value).toContain('sz="800"');
		expect(value).toContain('<a:srgbClr val="595959">');
		const category = element(xml, 'catAx');
		expect(category).toMatch(/<c:txPr><a:bodyPr[^>]* rot="-2700000"/u);
		expect(category).toMatch(/<c:spPr><a:ln><a:noFill/u);
		expect(category).toMatch(/<a:defRPr[^>]* i="1"/u);
	});

	it('writes an axis line width without a colour, keeping the automatic colour', async () => {
		const { xml, chart } = await save('bar', {
			categories: CATEGORIES,
			series: [{ name: 'A', values: [1, 2, 3] }],
			axes: { valAx: { line: { width: 1.5 } } },
		});
		const value = element(xml, 'valAx');
		expect(value).toContain('<c:spPr><a:ln w="19050"></a:ln></c:spPr>');
		expect(chart.axes?.find((axis) => axis.axisType === 'valAx')?.spPr).toStrictEqual({
			strokeWidth: 1.5,
		});
	});

	it('writes the title and legend fonts', async () => {
		const { xml, chart } = await save('bar', {
			categories: CATEGORIES,
			title: 'Sales',
			titleFont: { size: 20, bold: true, italic: true, color: '#1F4E79', family: 'Georgia' },
			legendFont: { size: 9, italic: true, family: 'Arial' },
			series: [{ name: 'A', values: [1, 2, 3] }],
		});
		const title = element(xml, 'title');
		expect(title).toMatch(/sz="2000"/u);
		expect(title).toMatch(/ b="1"/u);
		expect(title).toMatch(/ i="1"/u);
		expect(title).toContain('<a:srgbClr val="1F4E79">');
		expect(title).toContain('typeface="Georgia"');
		const legend = element(xml, 'legend');
		expect(legend).toMatch(/<a:defRPr[^>]* sz="900"/u);
		expect(legend).toMatch(/ i="1"/u);
		expect(legend).toContain('typeface="Arial"');
		expect(chart.style).toMatchObject({ titleFontSize: 20, titleFontBold: true });
	});
});

describe('declarative chart inputs for generated report charts', () => {
	it('writes one series label group with every flag, one line, centred and turned', async () => {
		const { xml } = await save('bar', {
			categories: CATEGORIES,
			series: [
				{
					name: 'A',
					values: [1, 2, 3],
					dataLabels: { showValue: true, wrap: false, anchor: 'ctr', rotation: -90 },
				},
			],
		});
		const labels = element(element(xml, 'ser'), 'dLbls');
		expect(labels).not.toContain('<c:dLbl>');
		for (const flag of [
			'showLegendKey',
			'showCatName',
			'showSerName',
			'showPercent',
			'showBubbleSize',
		])
			expect(labels).toContain(`<c:${flag} val="0">`);
		expect(labels).toContain('<c:showVal val="1">');
		const bodyPr = /<a:bodyPr[^>]*>/u.exec(labels)?.[0] ?? '';
		expect(bodyPr).toContain('wrap="none"');
		expect(bodyPr).toContain('anchor="ctr"');
		expect(bodyPr).toContain('rot="-5400000"');
	});

	it('writes a point with no fill as a:noFill', async () => {
		const { xml } = await save('doughnut', {
			categories: CATEGORIES,
			series: [{ name: 'D', values: [1, 2, 3], pointColors: ['#112233', 'none', null] }],
		});
		const points = elements(xml, 'dPt');
		expect(points).toHaveLength(2);
		expect(points[0]).toContain('<a:srgbClr val="112233">');
		expect(element(points[1]!, 'spPr')).toMatch(/^<c:spPr><a:noFill/u);
	});

	it('writes axis line dash, gridline lines, tick marks and crossing', async () => {
		const { xml, chart } = await save('bar', {
			categories: CATEGORIES,
			series: [{ name: 'A', values: [1, 2, 3] }],
			axes: {
				valAx: {
					line: { color: '#808080', width: 0.75, dash: 'sysDash' },
					majorGridlinesLine: { color: '#D9D9D9', width: 0.5, dash: 'dash' },
					majorTickMark: 'none',
					minorTickMark: 'none',
					crosses: 'max',
					crossBetween: 'midCat',
				},
				catAx: { majorTickMark: 'out', line: 'none' },
			},
		});
		const value = element(xml, 'valAx');
		const axisLine = /<c:spPr>((?!<\/c:spPr>).)*<\/c:spPr>/su.exec(
			value.replace(/<c:majorGridlines>.*?<\/c:majorGridlines>/su, ''),
		)![0];
		expect(axisLine).toContain('<a:ln w="9525">');
		expect(axisLine).toContain('<a:srgbClr val="808080">');
		expect(axisLine).toContain('<a:prstDash val="sysDash">');
		const grid = element(value, 'majorGridlines');
		expect(grid).toContain('<a:ln w="6350">');
		expect(grid).toContain('<a:srgbClr val="D9D9D9">');
		expect(grid).toContain('<a:prstDash val="dash">');
		expect(value).toContain('<c:majorTickMark val="none">');
		expect(value).toContain('<c:minorTickMark val="none">');
		expect(value).toContain('<c:crosses val="max">');
		expect(value).toContain('<c:crossBetween val="midCat">');
		const category = element(xml, 'catAx');
		expect(category).toContain('<c:majorTickMark val="out">');
		expect(category).toMatch(/<c:spPr><a:ln><a:noFill/u);
		const reloaded = chart.axes?.find((axis) => axis.axisType === 'valAx');
		expect(reloaded).toMatchObject({
			crosses: 'max',
			crossBetween: 'midCat',
			majorGridlines: true,
		});
	});

	it('writes title runs with their own fonts', async () => {
		const { xml, chart } = await save('bar', {
			categories: CATEGORIES,
			titleRuns: [
				{ text: 'Sales ', bold: true, size: 14 },
				{ text: '2026', italic: true, size: 10, color: '#C00000' },
			],
			titleFont: { family: 'Georgia' },
			series: [{ name: 'A', values: [1, 2, 3] }],
		});
		const title = element(xml, 'title');
		const runs = title.match(/<a:r>.*?<\/a:r>/gsu) ?? [];
		expect(runs).toHaveLength(2);
		expect(runs[0]).toMatch(/sz="1400"/u);
		expect(runs[0]).toMatch(/ b="1"/u);
		expect(runs[0]).toContain('<a:t>Sales </a:t>');
		expect(runs[1]).toMatch(/sz="1000"/u);
		expect(runs[1]).toMatch(/ i="1"/u);
		expect(runs[1]).toContain('<a:srgbClr val="C00000">');
		expect(title).toContain('typeface="Georgia"');
		expect(chart.titleRuns?.map((run) => run.text)).toStrictEqual(['Sales ', '2026']);
	});
});

describe('declarative chart input validation', () => {
	const input = (dataLabels: ChartInput['dataLabels'], grouping?: ChartInput['grouping']) => ({
		categories: CATEGORIES,
		...(grouping ? { grouping } : {}),
		dataLabels,
		series: [{ name: 'A', values: [1, 2, 3] }],
	});

	it('refuses a label position the chart type does not offer', () => {
		expect(() => createChartElement('bar', input({ position: 'outEnd' }, 'stacked'))).toThrow(
			RangeError,
		);
		expect(() => createChartElement('doughnut', input({ position: 'ctr' }))).toThrow(
			/takes no position/u,
		);
		expect(() => createChartElement('bar', input({ position: 'outEnd' }))).not.toThrow();
	});

	it('refuses fields the chart type has no element for', () => {
		const series = (extra: object) => ({
			categories: CATEGORIES,
			series: [{ name: 'A', values: [1, 2, 3], ...extra }],
		});
		expect(() => createChartElement('bar', series({ smooth: true }))).toThrow(/smooth/u);
		expect(() => createChartElement('line', series({ xValues: [1, 2] }))).toThrow(RangeError);
		expect(() => createChartElement('pie', series({ secondaryAxis: true }))).toThrow(
			/secondary axis/u,
		);
		expect(() => createChartElement('pie', series({ chartType: 'line' }))).toThrow(/combine/u);
		expect(() =>
			createChartElement('bar', { ...series({}), dataLabels: { rotation: 120 } }),
		).toThrow(RangeError);
		expect(() => createChartElement('bar', { ...series({}), palette: ['red'] })).toThrow(
			/hex colour/u,
		);
	});
});
