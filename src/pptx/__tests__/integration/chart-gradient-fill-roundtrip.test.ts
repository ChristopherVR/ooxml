import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import type { ChartGradientInput, ChartSeriesInput } from '../../core/builders/sdk';
import {
	PresentationBuilder,
	setChartDataPointGradient,
	setChartSeriesColor,
	setChartSeriesGradient,
	updateChartSeriesValues,
} from '../../core/builders/sdk';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxChartType } from '../../core/types/chart';
import type { ChartPptxElement } from '../../core/types/elements';

type Loaded = Awaited<ReturnType<PptxHandler['load']>>;

const CHART_PART = 'ppt/charts/chart1.xml';

const topToBottom: ChartGradientInput = {
	angle: 90,
	stops: [
		{ color: '#60A5FA', position: 0 },
		{ color: '#1E3A8A', position: 100 },
	],
};

function chartOf(data: Loaded): ChartPptxElement {
	return data.slides[0].elements.find((element) => element.type === 'chart') as ChartPptxElement;
}

async function readPart(bytes: Uint8Array, part = CHART_PART): Promise<string> {
	return (await JSZip.loadAsync(bytes)).file(part)!.async('string');
}

async function writePart(bytes: Uint8Array, xml: string): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(bytes);
	zip.file(CHART_PART, xml);
	return zip.generateAsync({ type: 'uint8array' });
}

/** Every `c:ser` element, in document order. */
function seriesXml(xml: string): string[] {
	return [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/gu)].map((match) => match[0]);
}

/** A series' own `c:spPr` (it precedes any `c:dPt` in every series schema). */
function seriesSpPr(ser: string): string | undefined {
	return /<c:spPr>[\s\S]*?<\/c:spPr>/u.exec(ser)?.[0];
}

async function createDeck(type: PptxChartType, series: ChartSeriesInput[]): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(
		createSlide('Blank')
			.addChart(
				type,
				{ categories: ['A', 'B', 'C'], series },
				{ x: 20, y: 20, width: 500, height: 300 },
			)
			.build(),
	);
	return handler.save(data.slides);
}

async function load(bytes: Uint8Array): Promise<{ handler: PptxHandler; data: Loaded }> {
	const handler = new PptxHandler();
	return { handler, data: await handler.load(bytes.buffer as ArrayBuffer) };
}

async function editAndSave(bytes: Uint8Array, edit: (chart: ChartPptxElement) => void) {
	const { handler, data } = await load(bytes);
	edit(chartOf(data));
	data.slides[0].isDirty = true;
	return handler.save(data.slides);
}

describe('chart gradient fills: new charts', () => {
	it('writes a series gradient into c:ser/c:spPr in OOXML units', async () => {
		const saved = await createDeck('bar', [
			{ name: 'Revenue', values: [1, 2, 3], color: '#FF0000', gradientFill: topToBottom },
		]);
		const spPr = seriesSpPr(seriesXml(await readPart(saved))[0]);
		expect(spPr).toBe(
			'<c:spPr><a:gradFill rotWithShape="1"><a:gsLst>' +
				'<a:gs pos="0"><a:srgbClr val="60A5FA"></a:srgbClr></a:gs>' +
				'<a:gs pos="100000"><a:srgbClr val="1E3A8A"></a:srgbClr></a:gs>' +
				'</a:gsLst><a:lin ang="5400000" scaled="0"></a:lin></a:gradFill></c:spPr>',
		);
	});

	it('writes a radial gradient with a:path circle and a fillToRect from the focal point', async () => {
		const saved = await createDeck('pie', [
			{
				name: 'Share',
				values: [1, 2, 3],
				gradientFill: { type: 'radial', focalPoint: { x: 0.3, y: 0.6 }, stops: topToBottom.stops },
			},
		]);
		const xml = await readPart(saved);
		expect(xml).toContain(
			'<a:path path="circle"><a:fillToRect l="30000" t="60000" r="70000" b="40000"></a:fillToRect></a:path>',
		);
		expect(xml).not.toContain('<a:lin ');
	});

	it('writes stop opacity as a:alpha', async () => {
		const saved = await createDeck('area', [
			{
				name: 'Fade',
				values: [1, 2, 3],
				gradientFill: {
					stops: [
						{ color: '#ED7D31', position: 0, opacity: 0.25 },
						{ color: '#ED7D31', position: 100 },
					],
				},
			},
		]);
		await expect(readPart(saved)).resolves.toContain(
			'<a:gs pos="0"><a:srgbClr val="ED7D31"><a:alpha val="25000"></a:alpha></a:srgbClr></a:gs>',
		);
	});

	it('round-trips create, save, load, save with a stable model and XML', async () => {
		const radial: ChartGradientInput = {
			type: 'radial',
			stops: [
				{ color: '#FFFFFF', position: 0, opacity: 0.5 },
				{ color: '#C00000', position: 100 },
			],
		};
		const saved = await createDeck('bar', [
			{ name: 'Linear', values: [1, 2, 3], gradientFill: topToBottom },
			{ name: 'Radial', values: [3, 2, 1], gradientFill: radial },
		]);
		const { data } = await load(saved);
		const [first, second] = chartOf(data).chartData!.series;
		expect(first.gradientFill).toStrictEqual({
			type: 'linear',
			angle: 90,
			stops: [
				{ color: '#60A5FA', position: 0 },
				{ color: '#1E3A8A', position: 100 },
			],
		});
		expect(second.gradientFill).toStrictEqual({
			type: 'radial',
			focalPoint: { x: 0.5, y: 0.5 },
			stops: [
				{ color: '#FFFFFF', position: 0, opacity: 0.5 },
				{ color: '#C00000', position: 100 },
			],
		});

		const resaved = await editAndSave(saved, (chart) => {
			updateChartSeriesValues(chart, 0, [4, 5, 6]);
		});
		expect(seriesXml(await readPart(resaved)).map(seriesSpPr)).toStrictEqual(
			seriesXml(await readPart(saved)).map(seriesSpPr),
		);
		const { data: reloaded } = await load(resaved);
		expect(chartOf(reloaded).chartData!.series.map((s) => s.gradientFill)).toStrictEqual([
			first.gradientFill,
			second.gradientFill,
		]);
	});
});

describe('chart gradient fills: existing charts', () => {
	it('sets and clears a series gradient on a loaded solid series', async () => {
		const saved = await createDeck('bar', [{ name: 'Solid', values: [1, 2, 3], color: '#4472C4' }]);

		const withGradient = await editAndSave(saved, (chart) => {
			setChartSeriesGradient(chart, 0, topToBottom);
		});
		const spPr = seriesSpPr(seriesXml(await readPart(withGradient))[0])!;
		expect(spPr).toContain('<a:gradFill');
		expect(spPr).not.toContain('<a:solidFill');
		const { data } = await load(withGradient);
		expect(chartOf(data).chartData!.series[0].gradientFill?.stops).toHaveLength(2);

		const cleared = await editAndSave(withGradient, (chart) => {
			setChartSeriesGradient(chart, 0, null);
		});
		const clearedSpPr = seriesSpPr(seriesXml(await readPart(cleared))[0])!;
		expect(clearedSpPr).not.toContain('<a:gradFill');
		const { data: reloaded } = await load(cleared);
		expect(chartOf(reloaded).chartData!.series[0].gradientFill).toBeUndefined();
	});

	it('lets the last colour/gradient call win on save', async () => {
		const saved = await createDeck('bar', [{ name: 'S', values: [1, 2, 3], color: '#4472C4' }]);
		const colourLast = await editAndSave(saved, (chart) => {
			setChartSeriesGradient(chart, 0, topToBottom);
			setChartSeriesColor(chart, 0, '#00B050');
		});
		const colourSpPr = seriesSpPr(seriesXml(await readPart(colourLast))[0])!;
		expect(colourSpPr).toContain('<a:srgbClr val="00B050"></a:srgbClr>');
		expect(colourSpPr).not.toContain('<a:gradFill');

		const gradientLast = await editAndSave(colourLast, (chart) => {
			setChartSeriesColor(chart, 0, '#FF0000');
			setChartSeriesGradient(chart, 0, topToBottom);
		});
		const gradientSpPr = seriesSpPr(seriesXml(await readPart(gradientLast))[0])!;
		expect(gradientSpPr).toContain('<a:gradFill');
		expect(gradientSpPr).not.toContain('<a:solidFill');
	});

	it('writes and removes a per-point gradient as c:dPt', async () => {
		const saved = await createDeck('bar', [{ name: 'S', values: [1, 2, 3], color: '#4472C4' }]);
		const withPoint = await editAndSave(saved, (chart) => {
			setChartDataPointGradient(chart, 0, 2, topToBottom);
		});
		const ser = seriesXml(await readPart(withPoint))[0];
		expect(ser).toMatch(
			/<c:dPt><c:idx val="2"><\/c:idx>[\s\S]*?<c:spPr><a:gradFill rotWithShape="1">[\s\S]*?<\/c:dPt>/u,
		);
		expect(ser.indexOf('<c:dPt>')).toBeLessThan(ser.indexOf('<c:cat>'));
		const { data } = await load(withPoint);
		expect(chartOf(data).chartData!.series[0].dataPoints).toStrictEqual([
			{
				idx: 2,
				gradientFill: {
					type: 'linear',
					angle: 90,
					stops: [
						{ color: '#60A5FA', position: 0 },
						{ color: '#1E3A8A', position: 100 },
					],
				},
			},
		]);

		const cleared = await editAndSave(withPoint, (chart) => {
			setChartDataPointGradient(chart, 0, 2, null);
		});
		expect(seriesXml(await readPart(cleared))[0]).not.toContain('<c:dPt>');
	});

	it('keeps an untouched themed gradient c:spPr byte-for-byte when only values change', async () => {
		const themed =
			'<c:spPr><a:gradFill flip="none" rotWithShape="1"><a:gsLst>' +
			'<a:gs pos="0"><a:schemeClr val="accent1"><a:lumMod val="60000"></a:lumMod><a:lumOff val="40000"></a:lumOff></a:schemeClr></a:gs>' +
			'<a:gs pos="100000"><a:schemeClr val="accent1"><a:shade val="76000"></a:shade></a:schemeClr></a:gs>' +
			'</a:gsLst><a:lin ang="5400000" scaled="1"></a:lin><a:tileRect></a:tileRect></a:gradFill>' +
			'<a:ln w="9525"><a:solidFill><a:schemeClr val="lt1"></a:schemeClr></a:solidFill></a:ln>' +
			'<a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">' +
			'<a:srgbClr val="000000"><a:alpha val="63000"></a:alpha></a:srgbClr></a:outerShdw></a:effectLst>' +
			'<a:scene3d><a:camera prst="orthographicFront"></a:camera><a:lightRig rig="threePt" dir="t"></a:lightRig></a:scene3d>' +
			'<a:sp3d><a:bevelT></a:bevelT></a:sp3d></c:spPr>';
		const generated = await createDeck('bar', [
			{ name: 'Themed', values: [1, 2, 3], color: '#4472C4' },
		]);
		const xml = await readPart(generated);
		const authored = await writePart(
			generated,
			xml.replace(seriesSpPr(seriesXml(xml)[0])!, themed),
		);
		const { data } = await load(authored);
		expect(chartOf(data).chartData!.series[0].gradientFill?.stops).toHaveLength(2);

		const resaved = await editAndSave(authored, (chart) => {
			updateChartSeriesValues(chart, 0, [7, 8, 9]);
		});
		const ser = seriesXml(await readPart(resaved))[0];
		expect(ser).toContain('<c:v>9</c:v>');
		expect(seriesSpPr(ser)).toBe(themed);
	});
});

describe('chart gradient fills: line-drawn families', () => {
	it('rejects a gradient on a line-drawn chart up front', async () => {
		await expect(
			createDeck('line', [{ name: 'L', values: [1, 2, 3], gradientFill: topToBottom }]),
		).rejects.toThrow(/area-filled chart types/u);

		const saved = await createDeck('scatter', [{ name: 'S', values: [1, 2, 3], color: '#4472C4' }]);
		const { data } = await load(saved);
		expect(() => setChartSeriesGradient(chartOf(data), 0, topToBottom)).toThrow(
			/drawn as "scatter"/u,
		);
		expect(() => setChartDataPointGradient(chartOf(data), 0, 1, topToBottom)).toThrow(
			/area-filled chart types/u,
		);
	});

	it('ignores a gradient set directly on a line series model when saving', async () => {
		const saved = await createDeck('line', [{ name: 'L', values: [1, 2, 3], color: '#4472C4' }]);
		const resaved = await editAndSave(saved, (chart) => {
			chart.chartData!.series[0].gradientFill = {
				type: 'linear',
				angle: 0,
				stops: [
					{ color: '#000000', position: 0 },
					{ color: '#FFFFFF', position: 100 },
				],
			};
		});
		const xml = await readPart(resaved);
		expect(xml).not.toContain('<a:gradFill');
		expect(seriesSpPr(seriesXml(xml)[0])).toContain('<a:ln><a:solidFill>');
	});
});
