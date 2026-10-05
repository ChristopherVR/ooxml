import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import type { ChartInput } from '../../core/builders/sdk';
import {
	PresentationBuilder,
	setChartAreaFormat,
	setChartAxis,
	setChartGroupOptions,
	updateChartSeriesValues,
} from '../../core/builders/sdk';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxChartType } from '../../core/types/chart';
import type { ChartPptxElement } from '../../core/types/elements';

type Loaded = Awaited<ReturnType<PptxHandler['load']>>;

const CHART_PART = 'ppt/charts/chart1.xml';

function chartOf(data: Loaded): ChartPptxElement {
	return data.slides[0].elements.find((element) => element.type === 'chart') as ChartPptxElement;
}

async function readPart(bytes: Uint8Array): Promise<string> {
	return (await JSZip.loadAsync(bytes)).file(CHART_PART)!.async('string');
}

async function writePart(bytes: Uint8Array, xml: string): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(bytes);
	zip.file(CHART_PART, xml);
	return zip.generateAsync({ type: 'uint8array' });
}

async function createDeck(
	type: PptxChartType,
	input: Partial<ChartInput> = {},
): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(
		createSlide('Blank')
			.addChart(
				type,
				{
					categories: ['A', 'B', 'C'],
					series: [
						{ name: 'S1', values: [1, 2, 3] },
						{ name: 'S2', values: [3, 2, 1] },
					],
					...input,
				},
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

/** The first element named `tag` (with its content), or undefined. */
function element(xml: string, tag: string): string | undefined {
	const re = new RegExp(`<c:${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</c:${tag}>)`, 'u');
	return re.exec(xml)?.[0];
}

function chartAreaSpPr(xml: string): string | undefined {
	return /<\/c:chart><c:spPr>[\s\S]*?<\/c:spPr>/u.exec(xml)?.[0];
}

const reportLike: Partial<ChartInput> = {
	barDirection: 'bar',
	gapWidth: 35,
	overlap: 10,
	chartArea: { fill: 'none', border: 'none', roundedCorners: false },
	plotArea: { fill: 'none', border: 'none' },
	axes: { catAx: { visible: false }, valAx: { visible: false, min: 0, max: 1 } },
};

describe('chart layout: new charts', () => {
	it('writes the model fields in schema order and reads them back', async () => {
		const saved = await createDeck('bar', reportLike);
		const xml = await readPart(saved);
		expect(xml).toMatch(
			/<c:chartSpace[^>]*><c:roundedCorners val="0"><\/c:roundedCorners><c:chart>/u,
		);
		expect(xml).toMatch(
			/<c:gapWidth val="35"><\/c:gapWidth><c:overlap val="10"><\/c:overlap><c:axId /u,
		);
		expect(xml.match(/<c:delete val="1">/gu)).toHaveLength(2);
		expect(xml).toMatch(
			/<c:scaling><c:orientation val="minMax"><\/c:orientation><c:max val="1"><\/c:max><c:min val="0"><\/c:min><\/c:scaling>/u,
		);
		expect(chartAreaSpPr(xml)).toBe(
			'</c:chart><c:spPr><a:noFill></a:noFill><a:ln><a:noFill></a:noFill></a:ln></c:spPr>',
		);

		const chartData = chartOf((await load(saved)).data).chartData!;
		expect(chartData).toMatchObject({ barGapWidth: 35, barOverlap: 10, roundedCorners: false });
		expect(chartData.style).toMatchObject({
			chartAreaFill: 'none',
			chartAreaBorder: 'none',
			plotAreaFill: 'none',
			plotAreaBorder: 'none',
		});
		expect(chartData.axes?.map((axis) => axis.deleted)).toStrictEqual([true, true]);
	});

	it('round-trips a doughnut slice angle and hole size, and keeps them on a values edit', async () => {
		const saved = await createDeck('doughnut', { firstSliceAngle: 90, holeSize: 30 });
		const chartData = chartOf((await load(saved)).data).chartData!;
		expect(chartData).toMatchObject({ firstSliceAngle: 90, doughnutHoleSize: 30 });
		const resaved = await editAndSave(saved, (chart) =>
			updateChartSeriesValues(chart, 0, [4, 5, 6]),
		);
		const xml = await readPart(resaved);
		expect(element(xml, 'firstSliceAng')).toBe('<c:firstSliceAng val="90"></c:firstSliceAng>');
		expect(element(xml, 'holeSize')).toBe('<c:holeSize val="30"></c:holeSize>');
	});
});

describe('chart layout: loaded charts', () => {
	it('applies setChart* edits on save and survives a reload', async () => {
		const saved = await createDeck('bar');
		const edited = await editAndSave(saved, (chart) => {
			setChartGroupOptions(chart, { gapWidth: 91, overlap: -4 });
			setChartAreaFormat(chart, 'chart', {
				fill: '#FFFFFF',
				border: 'none',
				roundedCorners: false,
			});
			setChartAxis(chart, 'catAx', { visible: false });
		});
		const xml = await readPart(edited);
		expect(element(xml, 'gapWidth')).toBe('<c:gapWidth val="91"></c:gapWidth>');
		expect(element(xml, 'overlap')).toBe('<c:overlap val="-4"></c:overlap>');
		expect(element(xml, 'roundedCorners')).toBe('<c:roundedCorners val="0"></c:roundedCorners>');
		expect(chartAreaSpPr(xml)).toBe(
			'</c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"></a:srgbClr></a:solidFill><a:ln><a:noFill></a:noFill></a:ln></c:spPr>',
		);
		const chartData = chartOf((await load(edited)).data).chartData!;
		expect(chartData).toMatchObject({ barGapWidth: 91, barOverlap: -4, roundedCorners: false });
		expect(chartData.axes?.find((axis) => axis.axisType === 'catAx')?.deleted).toBe(true);

		const cleared = await editAndSave(edited, (chart) => {
			setChartGroupOptions(chart, { overlap: null });
			setChartAreaFormat(chart, 'chart', null);
			setChartAxis(chart, 'catAx', { visible: true });
		});
		const clearedXml = await readPart(cleared);
		expect(element(clearedXml, 'overlap')).toBeUndefined();
		expect(element(clearedXml, 'roundedCorners')).toBeUndefined();
		expect(chartAreaSpPr(clearedXml)).toBeUndefined();
		expect(clearedXml.match(/<c:delete val="1">/gu)).toBeNull();
	});

	it('keeps an untouched authored layout byte-for-byte when only values change', async () => {
		const generated = await readPart(await createDeck('bar'));
		const themedSpPr =
			'<c:spPr><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="95000"></a:lumMod></a:schemeClr></a:solidFill>' +
			'<a:ln w="9525" cap="flat"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="25000"></a:lumMod></a:schemeClr></a:solidFill></a:ln>' +
			'<a:effectLst></a:effectLst></c:spPr>';
		const authoredXml = generated
			.replace(
				/(<c:chartSpace[^>]*>)/u,
				'$1<c:date1904 val="0"></c:date1904><c:roundedCorners val="0"></c:roundedCorners>',
			)
			.replace(
				'<c:gapWidth val="150"></c:gapWidth>',
				'<c:gapWidth val="35"></c:gapWidth><c:overlap val="10"></c:overlap>',
			)
			.replace('</c:chart>', `</c:chart>${themedSpPr}`)
			.replace('<c:delete val="0"></c:delete>', '<c:delete val="1"></c:delete>');
		const authored = await writePart(await createDeck('bar'), authoredXml);

		const resaved = await editAndSave(authored, (chart) =>
			updateChartSeriesValues(chart, 0, [7, 8, 9]),
		);
		const xml = await readPart(resaved);
		expect(xml).toContain('<c:v>9</c:v>');
		for (const tag of ['roundedCorners', 'gapWidth', 'overlap', 'catAx']) {
			expect(element(xml, tag)).toBe(element(authoredXml, tag));
		}
		expect(chartAreaSpPr(xml)).toBe(`</c:chart>${themedSpPr}`);
	});
});
