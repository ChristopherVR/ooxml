// Insert > Chart for extended charts in Word: a new `chartEx` part, its embedded workbook, the
// `mc:AlternateContent` drawing Word writes, and a reload that edits and saves the chart again.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { chartExFamilyOf, CHART_EX_FAMILIES, chartExSpaceDataSheet } from '../chart/index';
import { parseXml } from '../xml/index';
import {
	applyDocxChartExDataSheetEdits,
	chartsIn,
	createDocument,
	docxChartPaint,
	loadDocx,
	newChartImage,
	newDocxChartEx,
	saveDocx,
} from './index';
import type { TextRun } from './model';

const WORKBOOK = 'word/embeddings/Microsoft_Excel_Worksheet.xlsx';

async function insert(family: (typeof CHART_EX_FAMILIES)[number]) {
	const model = createDocument();
	const paragraph = model.blocks[0]!;
	if (paragraph.type !== 'paragraph') throw new Error('paragraph');
	paragraph.runs.push({
		text: '',
		image: newChartImage(newDocxChartEx(family, { title: 'New chart' })),
	} as TextRun);
	return saveDocx(model);
}

async function sheetCells(bytes: Uint8Array): Promise<Record<string, string>> {
	const workbook = await JSZip.loadAsync(bytes);
	const doc = parseXml(await workbook.file('xl/worksheets/sheet1.xml')!.async('string'));
	const cells: Record<string, string> = {};
	for (const cell of Array.from(doc.getElementsByTagName('c'))) {
		const value = cell.getElementsByTagName('v')[0] ?? cell.getElementsByTagName('t')[0];
		cells[cell.getAttribute('r')!] = value?.textContent ?? '';
	}
	return cells;
}

describe('inserting a Word extended chart', () => {
	it.each(CHART_EX_FAMILIES)('writes and reloads a new %s chart', async (family) => {
		const saved = await insert(family);
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
			'PartName="/word/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"',
		);
		expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
			'Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="charts/chartEx1.xml"',
		);
		const document = await zip.file('word/document.xml')!.async('string');
		expect(document).toMatch(/<mc:AlternateContent[^>]*><mc:Choice Requires="cx1">/);
		expect(document).toContain('uri="http://schemas.microsoft.com/office/drawing/2014/chartex"');
		const part = await zip.file('word/charts/chartEx1.xml')!.async('string');
		expect(part).toContain('<cx:externalData r:id="rId1" cx:autoUpdate="0"/>');
		expect(part).toMatch(/<cx:f>Sheet1!\$A\$2:\$[AB]\$\d+<\/cx:f>/);
		const chart = chartsIn((await loadDocx(saved)).model.blocks)[0]!;
		expect(chart.workbookPartName).toBe(WORKBOOK);
		expect(chartExFamilyOf(chart.chartExSpace!)).toBe(family);
		expect(chart.title).toBe('New chart');
		expect(docxChartPaint(chart, 480, 280).svg).toContain('<svg');
	});

	it('lays the data out in the workbook Edit Data opens', async () => {
		const zip = await JSZip.loadAsync(await insert('pareto'));
		const cells = await sheetCells(await zip.file(WORKBOOK)!.async('uint8array'));
		expect(cells).toMatchObject({ A2: 'Category 1', B2: '45', C1: 'Series 1' });
	});

	it.each(CHART_EX_FAMILIES)('saves a data edit of a reloaded %s chart', async (family) => {
		const loaded = await loadDocx(await insert(family));
		const model = structuredClone(loaded.model);
		const paragraph = model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('paragraph');
		const run = paragraph.runs.find((r) => 'image' in r && r.image?.chart) as TextRun & {
			image: { chart: NonNullable<TextRun['image']>['chart'] };
		};
		run.image.chart = applyDocxChartExDataSheetEdits(run.image.chart!, [
			{ kind: 'value', series: 0, category: 0, value: 99 },
		]);
		const chart = chartsIn((await loadDocx(await loaded.save(model))).model.blocks)[0]!;
		expect(chartExSpaceDataSheet(chart.chartExSpace!).series[0]!.values[0]).toBe(99);
	});

	it('saves a data edit of a reloaded pareto chart to its workbook', async () => {
		const loaded = await loadDocx(await insert('pareto'));
		const model = structuredClone(loaded.model);
		const paragraph = model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('paragraph');
		const run = paragraph.runs.find((r) => 'image' in r && r.image?.chart) as TextRun & {
			image: { chart: NonNullable<TextRun['image']>['chart'] };
		};
		run.image.chart = applyDocxChartExDataSheetEdits(run.image.chart!, [
			{ kind: 'value', series: 0, category: 4, value: 60 },
		]);
		const saved = await loaded.save(model);
		const chart = chartsIn((await loadDocx(saved)).model.blocks)[0]!;
		expect(chartExSpaceDataSheet(chart.chartExSpace!).series[0]!.values.at(-1)).toBe(60);
		const zip = await JSZip.loadAsync(saved);
		expect((await sheetCells(await zip.file(WORKBOOK)!.async('uint8array'))).B6).toBe('60');
	});
});
