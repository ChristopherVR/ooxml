import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { Schema } from 'prosemirror-model';
import { describe, expect, it } from 'vitest';
import { chartsIn, loadDocx, type DocxChart } from './index';
import type { TextRun } from './model';
import { at, expectParagraph } from './test-support/access';
import { imageNodeSpec } from './ui/inline-content-schema';
import { inlineNodeRun, runToInlineNodes } from './ui/run-adapter';
import { markSpecs } from './ui/schema-marks';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORKBOOK = Uint8Array.from([80, 75, 3, 4, 1, 2, 3, 4]);
const CHART_PARTS = ['word/charts/chart1.xml', 'word/charts/_rels/chart1.xml.rels'];

const chartDrawing = (relId: string, wrapper: 'inline' | 'anchor' = 'inline') =>
	`<w:drawing><wp:${wrapper}><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="${relId}"/></a:graphicData></a:graphic></wp:${wrapper}></w:drawing>`;

/**
 * A Word package around a real chart part: the Excel-authored `xl/charts/chart1.xml` of
 * `xlsx/__fixtures__/excel-features.xlsx`, stored as Word stores charts, with an embedded workbook.
 */
async function chartDocument(): Promise<Uint8Array> {
	const source = await JSZip.loadAsync(
		readFileSync(path.join(import.meta.dirname, '../xlsx/__fixtures__/excel-features.xlsx')),
	);
	const chartXml = (await source.file('xl/charts/chart1.xml')!.async('string')).replace(
		'</c:chartSpace>',
		'<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>',
	);
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>Intro</w:t></w:r></w:p><w:p><w:r>${chartDrawing('rId5')}</w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${REL}/chart" Target="charts/chart1.xml"/></Relationships>`,
	);
	zip.file('word/charts/chart1.xml', chartXml);
	zip.file(
		'word/charts/_rels/chart1.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>`,
	);
	zip.file('word/embeddings/Microsoft_Excel_Worksheet.xlsx', WORKBOOK);
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

async function texts(bytes: Uint8Array, names: string[]): Promise<string[]> {
	const zip = await JSZip.loadAsync(bytes);
	return Promise.all(names.map((name) => zip.file(name)!.async('string')));
}

describe('charts in a Word document', () => {
	it('loads the chart part into the drawing model', async () => {
		const loaded = await loadDocx(await chartDocument());
		const image = at(expectParagraph(loaded.model.blocks[1]).runs, 0).image;
		expect(image?.unsupported).toBe('Chart');
		const chart = image?.chart as DocxChart;
		expect(chart).toMatchObject({
			placement: 'inline',
			name: 'Chart 1',
			id: '1',
			relId: 'rId5',
			partName: 'word/charts/chart1.xml',
			title: 'Sales by region',
			workbookPartName: 'word/embeddings/Microsoft_Excel_Worksheet.xlsx',
		});
		expect(chart.notice).toContain('preserved unchanged');
		const group = chart.chartSpace?.plotArea.groups[0];
		expect(group).toMatchObject({ kind: 'bar', barDirection: 'col', grouping: 'clustered' });
		expect(group?.series.map((series) => series.tx?.text)).toEqual(['Sales', 'Cost']);
		// The Excel part's print settings are kept raw by the model, so nothing is reported.
		expect(chart.issues).toEqual([]);
		expect(chartsIn(loaded.model.blocks)).toEqual([chart]);
	});

	it('keeps the chart part, its relationships and workbook byte for byte through an edit', async () => {
		const original = await chartDocument();
		const loaded = await loadDocx(original);
		const edited = structuredClone(loaded.model);
		at(expectParagraph(edited.blocks[0]).runs, 0).text = 'Edited intro';
		const saved = await loaded.save(edited);
		expect(await texts(saved, CHART_PARTS)).toEqual(await texts(original, CHART_PARTS));
		const zip = await JSZip.loadAsync(saved);
		expect(
			await zip.file('word/embeddings/Microsoft_Excel_Worksheet.xlsx')!.async('uint8array'),
		).toEqual(WORKBOOK);
		const document = await zip.file('word/document.xml')!.async('string');
		expect(document).toContain('Edited intro');
		expect(document).toContain(`<c:chart xmlns:c="${C}" r:id="rId5"/>`);
		expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
			'Target="charts/chart1.xml"',
		);
		const reopened = await loadDocx(saved);
		expect(JSON.parse(JSON.stringify(chartsIn(reopened.model.blocks)))).toEqual(
			JSON.parse(JSON.stringify(chartsIn(loaded.model.blocks))),
		);
	});

	it('round-trips the chart through the editor run adapter unchanged', async () => {
		const loaded = await loadDocx(await chartDocument());
		const run = at(expectParagraph(loaded.model.blocks[1]).runs, 0) as TextRun;
		const schema = new Schema({
			nodes: {
				doc: { content: 'inline*' },
				text: { group: 'inline' },
				image: imageNodeSpec,
			},
			marks: markSpecs,
		});
		const [node] = runToInlineNodes(run, schema);
		expect(JSON.stringify(inlineNodeRun(node!))).toBe(JSON.stringify(run));
	});

	it('reports a chart whose relationship or part is missing', async () => {
		const zip = await JSZip.loadAsync(await chartDocument());
		zip.remove('word/charts/chart1.xml');
		const missingPart = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		expect(chartsIn(missingPart.model.blocks)[0]?.issues.map((issue) => issue.code)).toEqual([
			'CHART_PART_MISSING',
		]);
		const document = await zip.file('word/document.xml')!.async('string');
		zip.file('word/document.xml', document.replace('r:id="rId5"', 'r:id="rId9"'));
		const missingRel = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const chart = chartsIn(missingRel.model.blocks)[0];
		expect(chart?.partName).toBeUndefined();
		expect(chart?.issues.map((issue) => issue.code)).toEqual(['CHART_RELATIONSHIP_UNRESOLVED']);
	});
});
