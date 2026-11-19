import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	applyDocxChartDataSheetEdits,
	chartsIn,
	documentDrawingBlocks,
	loadDocx,
} from '../docx/index';
import type { DocumentModel } from '../docx/model';
import { PptxHandler } from '../pptx/core/PptxHandler';
import type { ChartPptxElement } from '../pptx/core/types/elements';
import { chartDataApplySheetEdits } from '../pptx/core/utils/chart-data-sheet';
import { getCell } from '../xlsx/cells';
import { readEmbeddedChartSpace, writeEmbeddedChartSpace } from '../xlsx/embedded-chart-space';
import { loadXlsx } from '../xlsx/read/index';
import { gridChartWorkbook } from '../xlsx/ui/dialogs/grid-chart';
import type { ChartDataSheetEdit } from './data-sheet';
import { applyChartSpaceDataSheetEdits } from './edit/data-sheet';
import { createChartWorkbookFromRows } from './workbook/create';

// A chart over whole worksheet columns, as Excel saves it, embedded with its workbook in each
// product. Column C holds a text cell, which the chart caches as 0.
const FULL = 1_048_575;
const LAST = FULL - 1;
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const COLUMNS = 'BCDEF';
const HEADS = ['One', 'Two', 'Three', 'Four', 'Five'];

const cached: Record<number, [number, number][]> = {
	0: [
		[0, 1],
		[1, 2],
		[2, 3],
	],
	1: [
		[0, 4],
		[1, 0],
		[2, 6],
	],
	2: [
		[0, 7],
		[1, 8],
		[2, 9],
		[LAST, 99],
	],
	3: [
		[0, 1],
		[1, 1],
		[2, 1],
	],
	4: [
		[0, 2],
		[1, 2],
		[2, 2],
		[3, 2],
	],
};

const series = (index: number) => {
	const column = COLUMNS[index]!;
	const pts = cached[index]!.map(([at, v]) => `<c:pt idx="${at}"><c:v>${v}</c:v></c:pt>`).join('');
	return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:strRef><c:f>Sheet1!$${column}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${HEADS[index]}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:val><c:numRef><c:f>Sheet1!$${column}$2:$${column}$1048576</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${FULL}"/>${pts}</c:numCache></c:numRef></c:val></c:ser>`;
};
const external = '<c:externalData r:id="rIdWb"><c:autoUpdate val="0"/></c:externalData>';
const chartXml = (data = external) =>
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${[0, 1, 2, 3, 4].map(series).join('')}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart>${data}</c:chartSpace>`;

function workbookRows() {
	const rows: (string | number | null)[][] = [];
	rows[0] = [null, ...HEADS];
	for (const [index, points] of Object.entries(cached))
		for (const [at, value] of points) (rows[at + 1] ??= [])[Number(index) + 1] = value;
	rows[2]![2] = 'not a number';
	return rows;
}

const zipBytes = (zip: JSZip) => zip.generateAsync({ type: 'uint8array' });
const fixture = (relative: string) => readFileSync(path.join(import.meta.dirname, relative));
const chartRels = `<Relationships xmlns="${PR}"><Relationship Id="rIdWb" Type="${R}/package" Target="../embeddings/Book1.xlsx"/></Relationships>`;

async function embed(zip: JSZip, folder: 'word' | 'ppt', name: string) {
	zip.file(`${folder}/charts/${name}`, chartXml());
	zip.file(`${folder}/charts/_rels/${name}.rels`, chartRels);
	zip.file(`${folder}/embeddings/Book1.xlsx`, await createChartWorkbookFromRows(workbookRows()));
	const types = await zip.file('[Content_Types].xml')!.async('string');
	if (!/Extension="xlsx"/.test(types))
		zip.file(
			'[Content_Types].xml',
			types.replace(
				/<Types[^>]*>/,
				(open) =>
					`${open}<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>`,
			),
		);
}

interface Saved {
	xml: string;
	book: Awaited<ReturnType<typeof loadXlsx>>;
}

async function saved(bytes: Uint8Array, folder: string, name: string): Promise<Saved> {
	const zip = await JSZip.loadAsync(bytes);
	return {
		xml: await zip.file(`${folder}/charts/${name}`)!.async('string'),
		book: await loadXlsx(await zip.file(`${folder}/embeddings/Book1.xlsx`)!.async('uint8array')),
	};
}

async function editInWord(edits: ChartDataSheetEdit[]): Promise<Saved> {
	const zip = await JSZip.loadAsync(fixture('../docx/__fixtures__/chart.docx'));
	await embed(zip, 'word', 'chart1.xml');
	const loaded = await loadDocx(await zipBytes(zip));
	const model = structuredClone(loaded.model) as DocumentModel;
	for (const block of documentDrawingBlocks(model))
		if (block.type === 'paragraph')
			for (const run of block.runs)
				if (run.image?.chart?.chartSpace)
					run.image.chart = applyDocxChartDataSheetEdits(run.image.chart, edits);
	expect(chartsIn(documentDrawingBlocks(model)).length).toBeGreaterThan(0);
	return saved(await loaded.save(model), 'word', 'chart1.xml');
}

async function editInPowerPoint(edits: ChartDataSheetEdit[]): Promise<Saved> {
	const zip = await JSZip.loadAsync(
		fixture('../pptx/__tests__/fixtures/embedded-assets-sample.pptx'),
	);
	await embed(zip, 'ppt', 'big.xml');
	const slide = Object.keys(zip.files).find((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))!;
	const rels = slide.replace('slides/', 'slides/_rels/') + '.rels';
	const change = async (name: string, edit: (xml: string) => string) =>
		zip.file(name, edit(await zip.file(name)!.async('string')));
	await change(slide, (xml) =>
		xml.replace(
			'</p:spTree>',
			`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9000" name="c"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rIdBig"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree>`,
		),
	);
	await change(rels, (xml) =>
		xml.replace(
			'</Relationships>',
			`<Relationship Id="rIdBig" Type="${R}/chart" Target="../charts/big.xml"/></Relationships>`,
		),
	);
	await change('[Content_Types].xml', (xml) =>
		xml.replace(
			'</Types>',
			'<Override PartName="/ppt/charts/big.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
		),
	);
	const bytes = await zipBytes(zip);
	const handler = new PptxHandler();
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	const chart = data.slides
		.flatMap((entry) => entry.elements)
		.find(
			(element): element is ChartPptxElement =>
				element.type === 'chart' && element.chartData?.chartPartPath === 'ppt/charts/big.xml',
		)!;
	chart.chartData = chartDataApplySheetEdits(chart.chartData!, edits)!;
	const out = new Uint8Array(await handler.save(data.slides));
	return saved(out, 'ppt', 'big.xml');
}

async function editInVisio(edits: ChartDataSheetEdit[]): Promise<Saved> {
	const zip = await JSZip.loadAsync(
		(await gridChartWorkbook(
			[
				[null, 'A'],
				['x', 1],
			],
			'column',
		))!,
	);
	const part = Object.keys(zip.files).find((name) => /^xl\/charts\/chart\d+\.xml$/.test(name))!;
	zip.file(part, chartXml(''));
	// The sheet keeps its drawing; only its cells become those of the large chart.
	const rows = await JSZip.loadAsync(await createChartWorkbookFromRows(workbookRows()));
	const sheetData = /<sheetData>.*<\/sheetData>/su.exec(
		await rows.file('xl/worksheets/sheet1.xml')!.async('string'),
	)![0];
	const sheetPart = 'xl/worksheets/sheet1.xml';
	zip.file(
		sheetPart,
		(await zip.file(sheetPart)!.async('string')).replace(/<sheetData>.*<\/sheetData>/su, sheetData),
	);
	const bytes = await zipBytes(zip);
	const { space } = (await readEmbeddedChartSpace(bytes))!;
	const out = await writeEmbeddedChartSpace(bytes, applyChartSpaceDataSheetEdits(space, edits));
	return {
		xml: await (await JSZip.loadAsync(out)).file(part)!.async('string'),
		book: await loadXlsx(out),
	};
}

const cell = ({ book }: Saved, row: number, col: number) =>
	getCell(book.sheets[0]!, row, col)?.value;
const counts = (xml: string) => [...xml.matchAll(/<c:ptCount val="(\d+)"/g)].map((m) => +m[1]!);
const hasPoint = (xml: string, index: number, value: number) => {
	const series = xml.split('<c:ser>')[index + 1]!;
	return series.includes(`idx="${value}"`);
};

describe.each([
	['Word', editInWord],
	['PowerPoint', editInPowerPoint],
	['Visio', editInVisio],
] as const)('%s edits of a whole-column chart', (_product, edit) => {
	it('changes a value of a series shorter than the sheet', async () => {
		const out = await edit([{ kind: 'value', series: 0, category: 1, value: 42 }]);
		expect(counts(out.xml).filter((count) => count === FULL)).toHaveLength(5);
		expect(hasPoint(out.xml, 2, LAST)).toBe(true);
		expect(out.xml).toContain('<c:v>42</c:v>');
		expect(cell(out, 2, 1)).toBe(42);
		expect(cell(out, 2, 2)).toBe('not a number');
		expect(cell(out, FULL, 3)).toBe(99);
	}, 60_000);

	it('leaves a text cell the chart caches as 0 alone when a series is renamed', async () => {
		const out = await edit([{ kind: 'series-name', series: 0, name: 'Renamed' }]);
		expect(cell(out, 0, 1)).toBe('Renamed');
		expect(cell(out, 2, 2)).toBe('not a number');
	}, 60_000);

	it('adds a category inside the declared range', async () => {
		const out = await edit([{ kind: 'add-category' }]);
		expect(counts(out.xml).filter((count) => count === FULL)).toHaveLength(5);
		expect(hasPoint(out.xml, 0, 4)).toBe(true);
		expect(hasPoint(out.xml, 2, LAST)).toBe(true);
	}, 60_000);

	it('removes a category and keeps the point in the last row of the range', async () => {
		const out = await edit([{ kind: 'remove-category', category: 0 }]);
		expect(counts(out.xml).filter((count) => count === FULL)).toHaveLength(5);
		expect(hasPoint(out.xml, 2, LAST - 1)).toBe(true);
		expect(hasPoint(out.xml, 2, LAST)).toBe(false);
	}, 60_000);
});

describe('series edits of a chart with a workbook', () => {
	it.each([
		['adds', { kind: 'add-series' }, 6],
		['removes', { kind: 'remove-series', series: 4 }, 4],
	] as const)(
		'%s a series of a whole-column chart in PowerPoint',
		async (_name, change, total) => {
			const out = await editInPowerPoint([change]);
			expect(out.xml.split('<c:ser>')).toHaveLength(total + 1);
			expect(hasPoint(out.xml, 2, LAST)).toBe(true);
		},
		60_000,
	);

	it('writes a fresh sheet for 200,000 rows without a dense copy', async () => {
		const { appendChartWorkbookSheet } = await import('./workbook/append-sheet');
		const rows: (string | number | null)[][] = [['Head']];
		for (let at = 1; at <= 200_000; at++) rows[at] = [at];
		const book = await createChartWorkbookFromRows([['x']]);
		const { bytes, sheetName } = await appendChartWorkbookSheet(book, () => rows);
		const loaded = await loadXlsx(bytes);
		const sheet = loaded.sheets.find((item) => item.name === sheetName)!;
		expect(getCell(sheet, 200_000, 0)?.value).toBe(200_000);
	}, 60_000);
});
