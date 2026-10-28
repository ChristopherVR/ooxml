// Local-only native comparison: provider cache stays in the output directory, never in fixtures.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import {
	parseChartExSpace,
	writeChartExSpaceOverSource,
	chartExSummary,
	renderChartSummarySvg,
} from '../chart';
const [nativeWorkbook, output] = process.argv.slice(2);
if (!nativeWorkbook || !output)
	throw new Error('Usage: bun chart-region-map-native-fixture.ts <native-map.xlsx> <output-dir>');
await mkdir(output, { recursive: true });
const native = await JSZip.loadAsync(await readFile(nativeWorkbook));
const nativeXml = await native.file('xl/charts/chartEx1.xml')!.async('string');
const nativeSpace = parseChartExSpace(nativeXml).chartExSpace;
const geography = nativeSpace.plotArea.series[0]?.layoutProperties?.geography;
if (!geography?.geoCacheXml) throw new Error('The native workbook has no geography cache.');
const zip = await JSZip.loadAsync(
	await readFile(new URL('../pptx/__tests__/fixtures/e2e/chart-gallery.pptx', import.meta.url)),
);
let part: string | undefined;
for (const name of Object.keys(zip.files)) {
	if (!/^ppt\/charts\/chart[^/]*\.xml$/u.test(name)) continue;
	const xml = await zip.file(name)!.async('string');
	if (xml.includes('layoutId="regionMap"')) {
		part = name;
		break;
	}
}
if (!part) throw new Error('The gallery has no region map.');
const original = await zip.file(part)!.async('string');
const source = parseChartExSpace(original).chartExSpace;
source.plotArea.series[0]!.layoutProperties ??= {};
source.plotArea.series[0]!.layoutProperties.geography = {
	...geography,
	projectionType: 'robinson',
	viewedRegionType: 'world',
};
// Literal geography data: do not refresh the gallery's Q1..Q4 workbook references on native open.
const countries = ['Australia', 'France', 'Canada', 'Brazil'];
source.chartData = {
	data: [
		{
			id: 0,
			dimensions: [
				{
					kind: 'strDim',
					type: 'cat',
					levels: [{ pointCount: 4, points: countries.map((value, index) => ({ index, value })) }],
				},
				{
					kind: 'numDim',
					type: 'colorVal',
					levels: [
						{
							pointCount: 4,
							points: [10, 30, 60, 90].map((value, index) => ({ index, value: String(value) })),
						},
					],
				},
			],
		},
	],
};
source.plotArea.series[0]!.dataId = 0;
delete source.plotArea.series[0]!.data;
const xml = writeChartExSpaceOverSource(original, source);
zip.file(part, xml);
let slide = 0;
for (const name of Object.keys(zip.files)) {
	const match = /^ppt\/slides\/_rels\/slide(\d+)\.xml\.rels$/u.exec(name);
	if (match && (await zip.file(name)!.async('string')).includes(path.posix.basename(part)))
		slide = Number(match[1]);
}
const sharedSvg = path.resolve(output, 'shared-region-map.svg');
await writeFile(
	sharedSvg,
	renderChartSummarySvg(chartExSummary(xml)!.summary!, { width: 840, height: 420 }),
);
const presentation = path.resolve(output, 'cached-region-map.pptx');
await writeFile(
	presentation,
	await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
);
await writeFile(
	path.resolve(output, 'manifest.json'),
	JSON.stringify(
		{
			presentation,
			cases: [{ id: 'region-map-cached', slide, shape: 2, width: 840, height: 420, sharedSvg }],
		},
		null,
		2,
	),
);
console.log(JSON.stringify({ presentation, slide, sharedSvg }));
