/**
 * A series colour and point colours, saved through the shared chart core and reloaded.
 *
 * `setChartSeriesColor` follows the Format pane (and `chart/edit` `setChartSeriesColor`): the
 * points lose their own fill, or their own line on a line-drawn series. The save used to apply
 * that rule only to the part, after the point edits: the model kept point colours the file had
 * lost, and a point coloured after its series in the same save was written as an empty
 * `<c:dPt><c:idx/><c:spPr/></c:dPt>`, which PowerPoint and LibreOffice draw in the automatic
 * colour. Now the model op clears the points, a point with nothing else of its own is removed so it
 * takes the series colour, and the save writes the points the model states, never an empty
 * `c:spPr`.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
	setChartDataPointFill,
	setChartSeriesColor,
} from '../../core/builders/sdk/chart-operations';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxChartDataPoint, PptxChartType } from '../../core/types';
import type { ChartPptxElement } from '../../core/types/elements';

type Slides = Awaited<ReturnType<PptxHandler['load']>>['slides'];

const chartOf = (slides: Slides): ChartPptxElement =>
	slides[0]!.elements.find((element) => element.type === 'chart') as ChartPptxElement;

const pointsOf = (xml: string, series = 0): string[] =>
	(xml.match(/<c:ser>.*?<\/c:ser>/gsu)?.[series] ?? '').match(/<c:dPt>.*?<\/c:dPt>/gsu) ?? [];

/**
 * A chart of `type` whose points `before` overrides, saved and loaded; `edit` changes the loaded
 * chart, which is saved again. Returns the second part, the edited model and the model reloaded
 * from the second part.
 */
async function editLoaded(
	type: PptxChartType,
	before: (element: ChartPptxElement) => void,
	edit: (element: ChartPptxElement) => void,
): Promise<{ xml: string; edited: ChartPptxElement; reloaded: ChartPptxElement }> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(
		createSlide('Blank')
			.addChart(
				type,
				{
					categories: ['a', 'b', 'c'],
					series: [
						{ name: 'A', values: [1, 2, 3] },
						{ name: 'B', values: [3, 2, 1] },
					],
				},
				{ x: 50, y: 50, width: 500, height: 300 },
			)
			.build(),
	);
	before(chartOf(data.slides));
	const loader = new PptxHandler();
	const loaded = await loader.load((await handler.save(data.slides)).buffer as ArrayBuffer);
	const edited = chartOf(loaded.slides);
	edit(edited);
	const bytes = await loader.save(loaded.slides);
	const xml = await (await JSZip.loadAsync(bytes)).file('ppt/charts/chart1.xml')!.async('string');
	const reloaded = chartOf((await new PptxHandler().load(bytes.buffer as ArrayBuffer)).slides);
	expect(xml).not.toMatch(EMPTY_POINT_SHAPE);
	return { xml, edited, reloaded };
}

const points = (element: ChartPptxElement, series = 0): PptxChartDataPoint[] =>
	element.chartData!.series[series]!.dataPoints ?? [];

/** PowerPoint and LibreOffice draw a point with an empty c:spPr in the automatic colour. */
const EMPTY_POINT_SHAPE = /<c:dPt>((?!<\/c:dPt>).)*<c:spPr\/>/su;

describe('series colour and point colours', () => {
	it.each(['bar', 'doughnut'] as const)(
		'clears the point fills of a recoloured %s in the model as in the part',
		async (type) => {
			const { xml, edited, reloaded } = await editLoaded(
				type,
				(element) => setChartDataPointFill(element, 0, 2, '#DBE4FF'),
				(element) => setChartSeriesColor(element, 0, '#3B5BDB'),
			);
			expect(xml).toContain('<a:srgbClr val="3B5BDB"/>');
			expect(xml).not.toContain('DBE4FF');
			// No override is left, so the point takes the series fill.
			expect(pointsOf(xml)).toStrictEqual([]);
			expect(points(edited)).toStrictEqual([]);
			expect(points(reloaded)).toStrictEqual([]);
		},
	);

	it.each(['bar', 'doughnut'] as const)(
		'keeps a %s point coloured after its series in the same save',
		async (type) => {
			const { xml, edited, reloaded } = await editLoaded(
				type,
				(element) => setChartDataPointFill(element, 0, 0, '#FFD43B'),
				(element) => {
					setChartSeriesColor(element, 0, '#3B5BDB');
					setChartDataPointFill(element, 0, 2, '#DBE4FF');
				},
			);
			expect(xml).toContain('<a:srgbClr val="3B5BDB"/>');
			const written = pointsOf(xml);
			// Point 0, cleared by the series colour, has no override left and takes the series fill.
			expect(written).toHaveLength(1);
			expect(written[0]).toContain(
				'<c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="DBE4FF"/></a:solidFill>',
			);
			expect(points(edited)).toStrictEqual([{ idx: 2, spPr: { fillColor: '#DBE4FF' } }]);
			expect(points(reloaded)).toStrictEqual(points(edited));
		},
	);

	it('clears the line of a line point but keeps its marker', async () => {
		const { xml, edited, reloaded } = await editLoaded(
			'line',
			(element) => {
				element.chartData!.series[0]!.dataPoints = [
					{ idx: 2, spPr: { strokeColor: '#DBE4FF' }, marker: { symbol: 'circle', size: 7 } },
				];
			},
			(element) => setChartSeriesColor(element, 0, '#3B5BDB'),
		);
		expect(xml).toMatch(/<c:ser>.*?<a:ln[^>]*><a:solidFill><a:srgbClr val="3B5BDB"\/>/su);
		expect(xml).not.toContain('DBE4FF');
		expect(pointsOf(xml)[0]).toContain('<c:symbol val="circle"/><c:size val="7"/>');
		expect(pointsOf(xml)[0]).not.toContain('<c:spPr');
		expect(points(edited)).toStrictEqual([{ idx: 2, marker: { symbol: 'circle', size: 7 } }]);
		expect(points(reloaded)).toStrictEqual(points(edited));
	});

	it('removes a point colour the model removes in the same save', async () => {
		const { xml, reloaded } = await editLoaded(
			'bar',
			(element) => setChartDataPointFill(element, 0, 1, '#DBE4FF'),
			(element) => {
				setChartDataPointFill(element, 0, 1, null);
				setChartSeriesColor(element, 0, '#3B5BDB');
			},
		);
		expect(xml).not.toContain('DBE4FF');
		expect(pointsOf(xml)).toStrictEqual([]);
		expect(points(reloaded)).toStrictEqual([]);
	});
});
