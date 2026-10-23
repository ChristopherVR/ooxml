import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	chartSpaceDataSheet,
	createChartWorkbook,
	DEFAULT_CHART_DATA_SHEET,
	setChartTitle,
	setChartType,
} from '../chart/index';
import { parseXml } from '../xml/index';
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
} from './index';
import type { TextRun } from './model';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORKBOOK = 'word/embeddings/Microsoft_Excel_Worksheet.xlsx';
const KEPT = ['word/charts/style1.xml', 'word/charts/colors1.xml', 'word/theme/theme1.xml'];

/** The Word chart fixture with an embedded workbook holding the chart's data, as Word stores it. */
async function chartDocument(): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(
		readFileSync(path.join(import.meta.dirname, '__fixtures__/chart.docx')),
	);
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
			`<Relationship Id="rId9" Type="${REL}/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>`,
		),
	);
	zip.file(WORKBOOK, await createChartWorkbook(chartSpaceDataSheet(space), 'Second'));
	return zip.generateAsync({ type: 'uint8array' });
}

/** Replaces the first chart of `model` (in place) with `edit(chart)`. */
function replaceChart(model: DocumentModel, edit: (chart: DocxChart) => DocxChart): void {
	for (const block of documentDrawingBlocks(model))
		if (block.type === 'paragraph')
			for (const run of block.runs)
				if (run.image?.chart) {
					run.image.chart = edit(run.image.chart);
					return;
				}
	throw new Error('No chart');
}

/** `ref -> value` of the first worksheet of a workbook part. */
async function workbookCells(zip: JSZip, part: string): Promise<Record<string, string>> {
	const workbook = await JSZip.loadAsync(await zip.file(part)!.async('uint8array'));
	const doc = parseXml(await workbook.file('xl/worksheets/sheet1.xml')!.async('string'));
	const cells: Record<string, string> = {};
	for (const cell of Array.from(doc.getElementsByTagName('c'))) {
		const value = cell.getElementsByTagName('v')[0] ?? cell.getElementsByTagName('t')[0];
		cells[cell.getAttribute('r')!] = value?.textContent ?? '';
	}
	return cells;
}

const read = (zip: JSZip, names: string[]) =>
	Promise.all(names.map((name) => zip.file(name)!.async('string')));

describe('saving edited Word charts', () => {
	it('saves an unedited document byte for byte', async () => {
		const original = await chartDocument();
		const loaded = await loadDocx(original);
		expect(await saveDocx(loaded.model)).toEqual(original);
	});

	it('writes edited data to the chart part and its workbook, and reloads it', async () => {
		const original = await chartDocument();
		const loaded = await loadDocx(original);
		const edited = structuredClone(loaded.model);
		replaceChart(edited, (chart) =>
			applyDocxChartDataSheetEdits(chart, [
				{ kind: 'value', series: 0, category: 2, value: 999 },
				{ kind: 'series-name', series: 1, name: 'Spend' },
				{ kind: 'add-category' },
			]),
		);
		const saved = await loaded.save(edited);
		const zip = await JSZip.loadAsync(saved);
		const before = await JSZip.loadAsync(original);
		expect(await read(zip, KEPT)).toEqual(await read(before, KEPT));
		const reopened = await loadDocx(saved);
		const chart = chartsIn(reopened.model.blocks)[0]!;
		const sheet = docxChartDataSheet(chart)!;
		expect(sheet.series.map((series) => series.name)).toEqual(['Sales', 'Spend']);
		expect(sheet.series[0]!.values[2]).toBe(999);
		expect(sheet.categories).toHaveLength(5);
		// The workbook Word's Edit Data opens holds what the chart caches.
		const cells = await workbookCells(zip, WORKBOOK);
		expect(cells.C1).toBe('Spend');
		expect(cells.B4).toBe('999');
		expect(cells.A6).toBe('Cat 5');
		sheet.categories.forEach((label, row) => expect(cells[`A${row + 2}`]).toBe(label));
		sheet.series.forEach((series, col) =>
			series.values.forEach((value, row) =>
				expect(cells[`${'BC'[col]}${row + 2}`]).toBe(String(value)),
			),
		);
	});

	it('writes a type and title change and keeps the drawing markup', async () => {
		const loaded = await loadDocx(await chartDocument());
		const edited = structuredClone(loaded.model);
		replaceChart(edited, (chart) =>
			editDocxChart(chart, (space) =>
				setChartTitle(setChartType(space, { kind: 'line' }), 'Trend'),
			),
		);
		const saved = await loaded.save(edited);
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file('word/charts/chart1.xml')!.async('string')).toContain('<c:lineChart>');
		const chart = chartsIn((await loadDocx(saved)).model.blocks)[0]!;
		expect(chart.title).toBe('Trend');
		expect(chart.chartSpace?.plotArea.groups[0]?.kind).toBe('line');
		expect(chart.relId).toBe(chartsIn(loaded.model.blocks)[0]!.relId);
	});
});

describe('inserting a Word chart', () => {
	it('creates the chart part, relationships, workbook and content types in a new document', async () => {
		const model = createDocument();
		const image = newChartImage(newDocxChart({ kind: 'pie' }, { title: 'Share' }));
		const paragraph = model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('paragraph');
		paragraph.runs.push({ text: '', image } as TextRun);
		const saved = await saveDocx(model);
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
			'PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"',
		);
		expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
			`Type="${REL}/chart" Target="charts/chart1.xml"`,
		);
		expect(await zip.file('word/document.xml')!.async('string')).toContain('<c:chart');
		const chart = chartsIn((await loadDocx(saved)).model.blocks)[0]!;
		expect(chart.title).toBe('Share');
		expect(chart.workbookPartName).toBe(WORKBOOK);
		expect(chart.chartSpace?.plotArea.groups[0]?.kind).toBe('pie');
		expect(docxChartDataSheet(chart)).toEqual(DEFAULT_CHART_DATA_SHEET);
		const cells = await workbookCells(zip, WORKBOOK);
		expect(cells).toMatchObject({ B1: 'Series 1', A5: 'Category 4', D5: '5' });
	});

	it('adds a second chart beside an existing one without touching it', async () => {
		const original = await chartDocument();
		const loaded = await loadDocx(original);
		const edited = structuredClone(loaded.model);
		const first = edited.blocks[0]!;
		if (first.type !== 'paragraph') throw new Error('paragraph');
		first.runs.push({ text: '', image: newChartImage(newDocxChart()) } as TextRun);
		const saved = await loaded.save(edited);
		const zip = await JSZip.loadAsync(saved);
		const before = await JSZip.loadAsync(original);
		expect(await read(zip, ['word/charts/chart1.xml'])).toEqual(
			await read(before, ['word/charts/chart1.xml']),
		);
		expect(zip.file('word/charts/chart2.xml')).not.toBeNull();
		expect(zip.file('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')).not.toBeNull();
		const charts = chartsIn((await loadDocx(saved)).model.blocks);
		expect(charts.map((chart) => chart.partName).sort()).toEqual([
			'word/charts/chart1.xml',
			'word/charts/chart2.xml',
		]);
	});
});
