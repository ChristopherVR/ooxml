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
import type { ChartPptxElement, PptxChartData } from '../pptx/core/types';

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
		slideIndex: number;
		slideCount: number;
		series: { name: string; values: number[] }[];
		categories: string[];
		cells: { sheet: string; address: string; value: string | number }[];
	}
> = {};
async function save(name: string, nativeType: number) {
	const expected = structuredClone(element.chartData!);
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
		throw new Error(`${name}: own-host save/reopen differs`);
	const zip = await JSZip.loadAsync(bytes);
	const neutral = parseChartSpace(await zip.file(back.chartPartPath!)!.async('string')).chartSpace;
	const cells = chartSpaceWorkbookWrites(neutral).flatMap((write) => {
		const range = parseChartFormulaRange(write.formula);
		const addresses = range ? expandRangeAddresses(range, write.values.length) : undefined;
		if (!range || !addresses)
			throw new Error('Unsupported native acceptance formula ' + write.formula);
		return addresses.map((address, index) => ({
			sheet: range.sheetName,
			address,
			value: write.isNumeric ? Number(write.values[index]) : write.values[index]!,
		}));
	});
	manifest[name] = {
		type: expected.chartType,
		nativeType,
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
writeFileSync(join(output, 'manifest.json'), JSON.stringify(manifest, null, 2));
console.log(`Wrote ${Object.keys(manifest).length} classic chart cases to ${output}`);
