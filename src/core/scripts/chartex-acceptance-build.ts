// Builds the files `chartex-acceptance.ps1` opens in Word, Excel and PowerPoint. Run with bun:
//   bun scripts/chartex-acceptance-build.ts <output folder>
// Each file holds one extended chart of every family Insert > Chart offers (pareto and region
// map included), created by the shared ChartEx core and saved by the product's own writer.
// `manifest.json` lists, per file, the families in document order and the first series' values
// Office must read back. Region maps carry no geography cache: Office fetches it from its own
// online service, which this script never contacts.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import {
	CHART_EX_FAMILIES,
	chartExFamilyOf,
	chartExSpaceDataSheet,
	type ChartExFamily,
	type ChartExSpace,
} from '../chart/index';
import {
	chartExSampleSheet,
	chartsIn,
	createDocument,
	documentDrawingBlocks,
	loadDocx,
	newChartImage,
	newDocxChartEx,
	saveDocx,
} from '../docx/index';
import type { TextRun } from '../docx/model';
import { createEditSession } from '../xlsx/edit/index';
import { xlsxChartExSpace } from '../xlsx/edit/chart-ex';
import { loadXlsx } from '../xlsx/read';
import { saveXlsx } from '../xlsx/write';
import { createWorkbook } from '../xlsx/workbook';
import { buildChartEx } from '../xlsx/ui/dialogs/insert-chart-ex-model';
import JSZip from 'jszip';
import { parseChartExSpace } from '../chart/index';
import { PresentationBuilder } from '../pptx/core/builders/sdk/PresentationBuilder';
import type { PptxChartSeries, PptxChartType } from '../pptx/core/types';

const out = process.argv[2];
if (!out)
	throw new Error('usage: bun scripts/chartex-acceptance-build.ts <output folder> [--verify]');
mkdirSync(out, { recursive: true });
const verifying = process.argv.includes('--verify');

interface Expectation {
	family: ChartExFamily;
	values: number[];
}
const manifest: Record<string, Expectation[]> = {};

const families = (process.env.CHARTEX_FAMILIES?.split(',') as ChartExFamily[] | undefined) ?? [
	...CHART_EX_FAMILIES,
];

async function word(): Promise<void> {
	const model = createDocument();
	model.blocks = families.map((family, index) => ({
		type: 'paragraph' as const,
		id: `chart-${index}`,
		runs: [
			{ text: '', image: newChartImage(newDocxChartEx(family, { title: family })) } as TextRun,
		],
	}));
	writeFileSync(join(out!, 'word-chartex.docx'), await saveDocx(model));
	manifest['word-chartex.docx'] = families.map((family) => ({
		family,
		values: chartExSampleSheet(family).series[0]!.values.map((v) => v ?? 0),
	}));
}

async function excel(): Promise<void> {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	const rows: (string | number)[][] = [
		['Group', 'Cause', 'Count', '', 'Country', 'Visitors'],
		['Shipping', 'Late', 42, '', 'France', 30],
		['Shipping', 'Damaged', 18, '', 'Brazil', 60],
		['Order', 'Wrong item', 27, '', 'Canada', 90],
		['Order', 'Missing', 9, '', 'India', 45],
	];
	rows.forEach((row, r) =>
		row.forEach((value, c) => value !== '' && session.setCellValue(0, r, c, value)),
	);
	const expected: Expectation[] = [];
	families.forEach((family, index) => {
		const map = family === 'regionMap';
		const range = map
			? { start: { row: 0, col: 4 }, end: { row: 4, col: 5 } }
			: family === 'treemap' || family === 'sunburst'
				? { start: { row: 0, col: 0 }, end: { row: 4, col: 2 } }
				: { start: { row: 0, col: 1 }, end: { row: 4, col: 2 } };
		const chart = buildChartEx(workbook, 0, range, family, { title: family });
		const top = 6 + index * 16;
		chart.anchor = {
			from: { row: top, col: 0, rowOffset: 0, colOffset: 0 },
			to: { row: top + 15, col: 8, rowOffset: 0, colOffset: 0 },
		};
		session.addChartEx(0, chart);
		expected.push({ family, values: map ? [30, 60, 90, 45] : [42, 18, 27, 9] });
	});
	writeFileSync(join(out!, 'excel-chartex.xlsx'), await saveXlsx(workbook));
	manifest['excel-chartex.xlsx'] = expected;
}

async function powerpoint(): Promise<void> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const types: Record<ChartExFamily, PptxChartType> = {
		waterfall: 'waterfall',
		clusteredColumn: 'histogram',
		pareto: 'histogram',
		boxWhisker: 'boxWhisker',
		treemap: 'treemap',
		sunburst: 'sunburst',
		funnel: 'funnel',
		regionMap: 'regionMap',
	};
	const expected: Expectation[] = [];
	for (const family of families) {
		const sample = chartExSampleSheet(family);
		const values = sample.series[0]!.values.map((v) => v ?? 0);
		const series: PptxChartSeries[] = [{ name: 'Series 1', values }];
		if (family === 'pareto') {
			series[0]!.histogramOptions = { layout: 'histogram', aggregateByCategory: true };
			series.push({ name: 'Cumulative %', values: [], histogramOptions: { layout: 'pareto' } });
		}
		data.slides.push(
			createSlide('Blank')
				.addChart(
					types[family],
					{
						categories: sample.categories,
						...('categoryLevels' in sample ? { categoryLevels: sample.categoryLevels } : {}),
						series,
					},
					{ x: 50, y: 50, width: 600, height: 380 },
				)
				.build(),
		);
		expected.push({ family, values });
	}
	writeFileSync(join(out!, 'powerpoint-chartex.pptx'), await handler.save(data.slides));
	manifest['powerpoint-chartex.pptx'] = expected;
}

/** The extended charts of a file Office saved again, in document order. */
async function officeSpaces(file: string): Promise<ChartExSpace[]> {
	const bytes = readFileSync(file);
	if (file.endsWith('.docx'))
		return chartsIn(documentDrawingBlocks((await loadDocx(bytes)).model)).flatMap((chart) =>
			chart.chartExSpace ? [chart.chartExSpace] : [],
		);
	if (file.endsWith('.xlsx')) {
		const book = await loadXlsx(bytes);
		return book.sheets.flatMap((sheet, s) =>
			sheet.drawings.flatMap((d, i) =>
				d.kind === 'unsupported' && d.chartExSpace ? [xlsxChartExSpace(book, s, i)] : [],
			),
		);
	}
	if (file.endsWith('.pptx')) {
		const zip = await JSZip.loadAsync(bytes);
		const parts = await Promise.all(
			Object.keys(zip.files)
				.filter((name) => /^ppt\/(charts|extendedCharts)\/chart(Ex)?\d+\.xml$/.test(name))
				.map(async (name) => [name, await zip.file(name)!.async('string')] as const),
		);
		// Slide order: PowerPoint numbers the parts it writes in slide order.
		return parts
			.filter(([, xml]) => xml.includes('cx:chartSpace'))
			.sort(([a], [b]) => a.localeCompare(b, 'en', { numeric: true }))
			.map(([, xml]) => parseChartExSpace(xml).chartExSpace);
	}
	throw new Error(`No reader for ${file}`);
}

async function verify(): Promise<void> {
	const expected = JSON.parse(readFileSync(join(out!, 'manifest.json'), 'utf8')) as typeof manifest;
	let failures = 0;
	for (const [name, charts] of Object.entries(expected)) {
		const saved = join(out!, name.replace(/\.(\w+)$/, '.office.$1'));
		if (!existsSync(saved)) {
			console.log(`${name}: no Office copy (${saved}); run chartex-acceptance.ps1 first`);
			failures++;
			continue;
		}
		const spaces = await officeSpaces(saved);
		console.log(`${name}: Office saved ${spaces.length} of ${charts.length} extended charts`);
		if (spaces.length !== charts.length) failures++;
		charts.forEach((want, index) => {
			const space = spaces[index];
			const family = space && chartExFamilyOf(space);
			const values = space ? chartExSpaceDataSheet(space).series[0]?.values.join(',') : '';
			const ok = family === want.family && values === want.values.join(',');
			if (!ok) failures++;
			console.log(
				`  #${index + 1} ${want.family}: Office kept ${family} [${values}] ${ok ? 'ok' : 'FAIL'}`,
			);
		});
	}
	if (failures) {
		console.log(`${failures} failure(s)`);
		process.exit(1);
	}
	console.log('all Office copies match');
}

if (verifying) await verify();
else {
	await word();
	await excel();
	await powerpoint();
	writeFileSync(join(out, 'manifest.json'), JSON.stringify(manifest, null, '\t'));
	console.log(`wrote ${Object.keys(manifest).length} files to ${out}`);
}
