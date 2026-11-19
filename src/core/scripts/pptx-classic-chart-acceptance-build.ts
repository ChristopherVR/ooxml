import JSZip from 'jszip';
import { parseChartSpace } from '../chart/index';
import {
	chartSpaceWorkbookWrites,
	parseChartFormulaRange,
	expandRangeAddresses,
} from '../chart/workbook/index';
// Generates repeated classic chart edits for native PowerPoint/Edit Data acceptance.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { PptxHandler } from '../pptx/core/PptxHandler';
import { setChartSeriesChartType } from '../pptx/core/builders/sdk/chart-operations';
import type { ChartPptxElement, PptxChartData } from '../pptx/core/types';
import { planWithLeftovers } from '../pptx/core/utils/chart-neutral-save-plan';

const output = process.argv[2];
if (!output) throw new Error('Expected an output directory.');
mkdirSync(output, { recursive: true });
const source = new Uint8Array(
	readFileSync(new URL('../pptx/__tests__/fixtures/e2e/chart-data-fidelity.pptx', import.meta.url)),
);
const handler = new PptxHandler();
const deck = await handler.load(source);
const slideIndex = deck.slides.findIndex((slide) =>
	slide.elements.some((entry) => entry.type === 'chart' && entry.chartData?.chartType === 'bar'),
);
if (slideIndex < 0) throw new Error('No native column chart found.');
const element = deck.slides[slideIndex]!.elements.find(
	(item) => item.type === 'chart',
) as ChartPptxElement;
const manifest: Record<
	string,
	{
		type: string;
		nativeType: number;
		/** Per-series native types of a combination chart (`Series.ChartType`), when checked. */
		nativeTypes?: number[];
		/** Per-series axis groups (`Series.AxisGroup`: 1 primary, 2 secondary), when checked. */
		axisGroups?: number[];
		/** The number of chart groups (`Chart.ChartGroups().Count`), when checked. */
		groups?: number;
		/** Which writer saved the part. */
		writer: string;
		slideIndex: number;
		slideCount: number;
		series: { name: string; values: number[] }[];
		categories: string[];
		cells: { sheet: string; address: string; value: string | number }[];
	}
> = {};
async function save(
	name: string,
	nativeType: number,
	nativeTypes?: number[],
	axisGroups?: number[],
	groups?: number,
) {
	const expected = structuredClone(element.chartData!);
	// Which writer saves the part: the neutral plan, or the object-tree writers for unclaimed fields.
	const baseline = (
		handler as unknown as { runtime: { chartDataBaselines: Map<string, string> } }
	).runtime.chartDataBaselines.get(expected.chartPartPath!);
	const unclaimed = baseline
		? planWithLeftovers(JSON.parse(baseline) as PptxChartData, expected).unclaimed
		: ['(no baseline)'];
	const bytes = await handler.save(deck.slides);
	const reopened = await new PptxHandler().load(bytes);
	const back = (
		reopened.slides[slideIndex]!.elements.find((item) => item.type === 'chart') as ChartPptxElement
	).chartData!;
	if (
		back.chartType !== expected.chartType ||
		JSON.stringify(back.categories) !== JSON.stringify(expected.categories) ||
		JSON.stringify(back.series.map(({ name, values }) => ({ name, values }))) !==
			JSON.stringify(expected.series.map(({ name, values }) => ({ name, values })))
	)
		throw new Error(
			`${name}: own-host save/reopen differs: ${JSON.stringify({
				type: [back.chartType, expected.chartType],
				categories: [back.categories, expected.categories],
				series: [back.series, expected.series].map((list) =>
					list.map(({ name, values }) => ({ name, values })),
				),
			})}`,
		);
	const zip = await JSZip.loadAsync(bytes);
	const neutral = parseChartSpace(await zip.file(back.chartPartPath!)!.async('string')).chartSpace;
	const cells = chartSpaceWorkbookWrites(neutral).flatMap((write) => {
		const range = parseChartFormulaRange(write.formula);
		const addresses = range ? expandRangeAddresses(range, write.cellCount) : undefined;
		if (!range || !addresses)
			throw new Error('Unsupported native acceptance formula ' + write.formula);
		return write.cells.map(({ index, value }) => ({
			sheet: range.sheetName,
			address: addresses[index]!,
			value: write.isNumeric ? Number(value) : value,
		}));
	});
	manifest[name] = {
		type: expected.chartType,
		nativeType,
		...(nativeTypes ? { nativeTypes } : {}),
		...(axisGroups ? { axisGroups } : {}),
		...(groups === undefined ? {} : { groups }),
		writer: unclaimed.length === 0 ? 'neutral' : `object-tree (${unclaimed.join(', ')})`,
		slideIndex: slideIndex + 1,
		slideCount: deck.slides.length,
		series: expected.series.map(({ name, values }) => ({ name, values })),
		categories: expected.categories,
		cells,
	};
	writeFileSync(join(output, name), bytes);
}
const chart = (): PptxChartData => element.chartData!;
chart().categories = ['Plan', 'Design', 'Build', 'Review', 'Ship'];
chart().series = [
	{ name: 'Primary growth', values: [2, 4, 6, 8, 10] },
	{ name: 'Second growth', values: [3, 6, 9, 12, 15] },
	{ name: 'Added growth', values: [1, 5, 7, 11, 13] },
];
chart().title = 'Classic growth';
await save('classic-growth.pptx', 51);
chart().chartType = 'pie';
chart().series = [chart().series[0]!];
chart().firstSliceAngle = 45;
await save('classic-pie.pptx', 5);
chart().chartType = 'line3D';
chart().series.push({ name: 'Line comparison', values: [5, 10, 15, 20, 25] });
await save('classic-line3d.pptx', -4101);
chart().series[0]!.name = 'Repeated line edit';
chart().series[0]!.values = [10, 20, 30, 40, 50];
await save('classic-repeated.pptx', -4101);
chart().chartType = 'bar';
chart().categories = ['Start', 'Finish'];
chart().series = [{ name: 'Shrunk', values: [7, 14] }];
chart().title = 'Shrunk chart';
await save('classic-shrunk.pptx', 51);
// A column chart whose last series becomes XY (Scatter): the scatter group shares the column axes.
chart().categories = ['Plan', 'Design', 'Build', 'Review', 'Ship'];
chart().series = [
	{ name: 'Columns A', values: [2, 4, 6, 8, 10] },
	{ name: 'Columns B', values: [3, 6, 9, 12, 15] },
	{ name: 'Scatter C', values: [1, 5, 7, 11, 13] },
];
chart().title = 'Column and scatter';
await save('classic-xy-combo.pptx', 51);
setChartSeriesChartType(element, 2, 'scatter');
await save('classic-xy-combo-scatter.pptx', 51, [51, 51, -4169]);
// Group subtype settings on kept groups: a 3-D column's bar shape and a radar's fill.
for (const series of chart().series) delete series.seriesChartType;
chart().chartType = 'bar3D';
await save('classic-bar3d.pptx', 54);
chart().barShape = 'cylinder';
chart().gapDepth = 220;
await save('classic-bar3d-cylinder.pptx', 92);
chart().chartType = 'radar';
delete chart().barShape;
delete chart().gapDepth;
await save('classic-radar.pptx', 81);
chart().radarStyle = 'filled';
await save('classic-radar-filled.pptx', 82);
// Wireframe, pie-of-pie and bubble edits on kept groups.
delete chart().radarStyle;
chart().chartType = 'surface';
await save('classic-surface.pptx', 83);
chart().wireframe = true;
await save('classic-surface-wireframe.pptx', 84);
delete chart().wireframe;
chart().chartType = 'ofPie';
chart().series = [{ name: 'Split', values: [8, 6, 4, 2, 1] }];
await save('classic-of-pie.pptx', 68);
chart().ofPieOptions = {
	...chart().ofPieOptions!,
	splitType: 'pos',
	splitPos: 2,
	secondPieSize: 60,
	gapWidth: 120,
};
await save('classic-of-pie-split.pptx', 68);
delete chart().ofPieOptions;
// Per-series combinations Excel 16 saves through `Series.ChartType`: a line series of horizontal
// bars on a secondary axis pair, and a series converted back rejoining its column group.
chart().chartType = 'bar';
chart().barDirection = 'bar';
chart().series = [
	{ name: 'Bars A', values: [2, 4, 6, 8, 10] },
	{ name: 'Bars B', values: [3, 6, 9, 12, 15] },
	{ name: 'Line C', values: [1, 5, 7, 11, 13] },
];
chart().title = 'Bars and line';
await save('classic-hbar.pptx', 57);
setChartSeriesChartType(element, 2, 'line');
await save('classic-hbar-line.pptx', 57, [57, 57, 65], [1, 1, 2]);
for (const series of chart().series) delete series.seriesChartType;
chart().chartType = 'bar';
await save('classic-hbar-back.pptx', 57, undefined, [1, 1, 1], 1);
chart().barDirection = 'col';
await save('classic-columns.pptx', 51);
setChartSeriesChartType(element, 1, 'line');
await save('classic-columns-line.pptx', 51, [51, 65, 51]);
for (const series of chart().series) delete series.seriesChartType;
chart().chartType = 'bar';
await save('classic-columns-rejoined.pptx', 51, undefined, undefined, 1);
// An XY chart (two value axes) with one line series, and bubble options.
chart().categories = ['1', '2', '3', '4', '5'];
await save('classic-columns-numbered.pptx', 51);
chart().chartType = 'scatter';
chart().title = 'XY and line';
await save('classic-xy.pptx', 74);
setChartSeriesChartType(element, 2, 'line');
await save('classic-xy-line.pptx', -4169, [-4169, -4169, 4]);
for (const series of chart().series) delete series.seriesChartType;
chart().chartType = 'bubble';
chart().series = chart().series.slice(0, 2);
await save('classic-bubble.pptx', 15);
chart().bubbleOptions = { ...chart().bubbleOptions, bubbleScale: 55, sizeRepresents: 'w' };
await save('classic-bubble-options.pptx', 15);
writeFileSync(join(output, 'manifest.json'), JSON.stringify(manifest, null, 2));
for (const [name, entry] of Object.entries(manifest)) console.log(`${name}: ${entry.writer}`);
console.log(`Wrote ${Object.keys(manifest).length} classic chart cases to ${output}`);
