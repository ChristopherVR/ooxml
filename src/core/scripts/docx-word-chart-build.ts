// Builds the documents `docx-word-chart-acceptance.ps1` opens in Word. Run with bun:
//   bun scripts/docx-word-chart-build.ts <output folder>
// Every file is saved by `saveDocx` after chart edits made with the shared chart core: a chart
// Word itself authored (`word-source.docx`, written by the PowerShell script through Word's own
// Insert Chart when Word is available), the committed Excel-authored chart fixture with an
// embedded workbook, and new charts inserted into a blank document and into Word's document.
// `manifest.json` lists, per file and per chart in document order, the series names and values
// Word must read back and the cells the embedded workbook must hold.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import {
	chartSpaceDataSheet,
	createChartWorkbook,
	setChartTitle,
	setChartType,
	type ChartDataSheet,
	type ChartDataSheetEdit,
} from '../chart/index';
import {
	applyDocxChartDataSheetEdits,
	chartsIn,
	createDocument,
	docxChartDataSheet,
	documentDrawingBlocks,
	editDocxChart,
	loadDocx,
	newChartImage,
	newDocxChart,
	saveDocx,
	type DocumentModel,
	type DocxChart,
} from '../docx/index';
import type { TextRun } from '../docx/model';

const out = process.argv[2];
if (!out) throw new Error('usage: bun scripts/docx-word-chart-build.ts <output folder>');
mkdirSync(out, { recursive: true });

interface ChartExpectation {
	kind: string;
	series: { name: string; values: (number | null)[] }[];
	/** `A1` reference -> text the embedded workbook's first sheet must show. */
	cells: Record<string, string>;
}
const manifest: Record<string, ChartExpectation[]> = {};

const columnLetter = (index: number) => String.fromCharCode(65 + index);

function expectation(chart: DocxChart): ChartExpectation {
	const sheet = docxChartDataSheet(chart) as ChartDataSheet;
	const cells: Record<string, string> = {};
	sheet.series.forEach((series, s) => {
		cells[`${columnLetter(s + 1)}1`] = series.name;
		series.values.forEach((value, c) => {
			if (value !== null) cells[`${columnLetter(s + 1)}${c + 2}`] = String(value);
		});
	});
	sheet.categories.forEach((label, c) => (cells[`A${c + 2}`] = label));
	return {
		kind: chart.chartSpace?.plotArea.groups[0]?.kind ?? '',
		series: sheet.series.map((series) => ({ name: series.name, values: [...series.values] })),
		cells,
	};
}

function replaceCharts(model: DocumentModel, edit: (chart: DocxChart) => DocxChart): void {
	for (const block of documentDrawingBlocks(model))
		if (block.type === 'paragraph')
			for (const run of block.runs)
				if (run.image?.chart?.chartSpace && !run.image.chart.layoutId)
					run.image.chart = edit(run.image.chart);
}

function appendChart(model: DocumentModel, chart: DocxChart): void {
	model.blocks.push({
		type: 'paragraph',
		id: `chart-${model.blocks.length}`,
		runs: [{ text: '', image: newChartImage(chart) } as TextRun],
	});
}

async function write(name: string, bytes: Uint8Array): Promise<void> {
	writeFileSync(join(out!, name), bytes);
	const reloaded = await loadDocx(bytes);
	manifest[name] = chartsIn(documentDrawingBlocks(reloaded.model))
		.filter((chart) => chart.chartSpace && !chart.layoutId)
		.map(expectation);
}

async function edited(source: Uint8Array, name: string, edit: (chart: DocxChart) => DocxChart) {
	const loaded = await loadDocx(source);
	const model = structuredClone(loaded.model);
	replaceCharts(model, edit);
	await write(name, await loaded.save(model));
}

const data = (edits: ChartDataSheetEdit[]) => (chart: DocxChart) =>
	applyDocxChartDataSheetEdits(chart, edits);

/** The Excel-authored chart fixture, stored as Word stores charts, with an embedded workbook. */
async function fixtureWithWorkbook(): Promise<Uint8Array> {
	const path = join(import.meta.dir, '..', 'docx', '__fixtures__', 'chart.docx');
	const zip = await JSZip.loadAsync(readFileSync(path));
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const space = chartsIn(loaded.model.blocks)[0]!.chartSpace!;
	const part = 'word/charts/chart1.xml';
	zip.file(
		part,
		(await zip.file(part)!.async('string')).replace(
			'</c:chartSpace>',
			'<c:externalData r:id="rId9"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>',
		),
	);
	const rels = 'word/charts/_rels/chart1.xml.rels';
	zip.file(
		rels,
		(await zip.file(rels)!.async('string')).replace(
			'</Relationships>',
			'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>',
		),
	);
	zip.file(
		'word/embeddings/Microsoft_Excel_Worksheet.xlsx',
		await createChartWorkbook(chartSpaceDataSheet(space), 'Second'),
	);
	const types = '[Content_Types].xml';
	const typesXml = await zip.file(types)!.async('string');
	if (!typesXml.includes('Extension="xlsx"'))
		zip.file(
			types,
			typesXml.replace(
				'<Default ',
				'<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Default ',
			),
		);
	return zip.generateAsync({ type: 'uint8array' });
}

const structural: ChartDataSheetEdit[] = [
	{ kind: 'value', series: 0, category: 1, value: 7.25 },
	{ kind: 'series-name', series: 1, name: 'Renamed' },
	{ kind: 'category-label', category: 0, label: 'First' },
	{ kind: 'add-series' },
	{ kind: 'add-category' },
	{ kind: 'value', series: 0, category: 4, value: 9 },
];

const fixture = await fixtureWithWorkbook();
await edited(
	fixture,
	'fixture-value.docx',
	data([{ kind: 'value', series: 0, category: 0, value: 42 }]),
);
await edited(fixture, 'fixture-structural.docx', data(structural));
await edited(fixture, 'fixture-type-line.docx', (chart) =>
	editDocxChart(chart, (space) => setChartTitle(setChartType(space, { kind: 'line' }), 'Lines')),
);

const blank = createDocument();
appendChart(blank, newDocxChart(undefined, { title: 'New column chart' }));
appendChart(blank, newDocxChart({ kind: 'pie' }, { title: 'New pie chart' }));
appendChart(blank, newDocxChart({ kind: 'bar', barDirection: 'bar', grouping: 'stacked' }));
await write('new-charts.docx', await saveDocx(blank));

const source = join(out, 'word-source.docx');
if (existsSync(source)) {
	const bytes = new Uint8Array(readFileSync(source));
	await edited(
		bytes,
		'word-value.docx',
		data([{ kind: 'value', series: 0, category: 0, value: 42 }]),
	);
	await edited(bytes, 'word-structural.docx', data(structural));
	await edited(bytes, 'word-type-line.docx', (chart) =>
		editDocxChart(chart, (space) => setChartType(space, { kind: 'line' })),
	);
	const loaded = await loadDocx(bytes);
	const model = structuredClone(loaded.model);
	appendChart(model, newDocxChart({ kind: 'doughnut' }, { title: 'Added beside Word chart' }));
	await write('word-added-chart.docx', await loaded.save(model));
}

writeFileSync(join(out, 'manifest.json'), JSON.stringify(manifest, null, 2));
console.log(`Wrote ${Object.keys(manifest).length} documents to ${out}`);
