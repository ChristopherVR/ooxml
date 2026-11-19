import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { chartsIn, loadDocx } from '../docx/index';
import type { DocumentModel } from '../docx/model';
import { documentDrawingBlocks } from '../docx/chart';
import { PptxHandler } from '../pptx/core/PptxHandler';
import type { ChartPptxElement } from '../pptx/core/types/elements';
import { loadXlsx } from '../xlsx/read/index';
import { chartCacheValues } from './data-cache';
import { MAX_CACHE_BLANK_POINTS } from './point-limit';

// A declared point count costs a few bytes, so one package of many small chart parts must not
// expand each of them: every format's loader shares one budget across its charts.
const COUNT = 8;
const FULL = 4;
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

const chartPart = `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>S!$A:$A</c:f><c:numCache><c:ptCount val="${MAX_CACHE_BLANK_POINTS}"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;
const indices = Array.from({ length: COUNT }, (_, index) => index);
const rels = (target: (index: number) => string) =>
	`<Relationships xmlns="${PR}">${indices.map((i) => `<Relationship Id="rIdC${i}" Type="${R}/chart" Target="${target(i)}"/>`).join('')}</Relationships>`;
const overrides = (folder: string) =>
	indices
		.map((i) => `<Override PartName="/${folder}/c${i}.xml" ContentType="${CHART_TYPE}"/>`)
		.join('');
const zipBytes = (zip: JSZip) => zip.generateAsync({ type: 'uint8array' });
const expectBudgeted = (lengths: number[]) => {
	expect(lengths).toHaveLength(COUNT);
	expect(lengths.filter((length) => length === MAX_CACHE_BLANK_POINTS)).toHaveLength(FULL);
	expect(lengths.filter((length) => length === 1)).toHaveLength(COUNT - FULL);
};

describe('the blank-point budget of a package', () => {
	it('is shared by the charts of a workbook', async () => {
		const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
		const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
		const anchor = (i: number) =>
			`<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="c${i}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rIdC${i}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
		const zip = new JSZip();
		zip.file(
			'[Content_Types].xml',
			`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>${overrides('xl/charts')}</Types>`,
		);
		zip.file(
			'_rels/.rels',
			`<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
		);
		zip.file(
			'xl/workbook.xml',
			`<workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
		);
		zip.file(
			'xl/_rels/workbook.xml.rels',
			`<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
		);
		zip.file(
			'xl/worksheets/sheet1.xml',
			`<worksheet xmlns="${S}" xmlns:r="${R}"><sheetData/><drawing r:id="rId1"/></worksheet>`,
		);
		zip.file(
			'xl/worksheets/_rels/sheet1.xml.rels',
			`<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/drawing" Target="../drawings/d.xml"/></Relationships>`,
		);
		zip.file(
			'xl/drawings/d.xml',
			`<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}">${indices.map(anchor).join('')}</xdr:wsDr>`,
		);
		zip.file(
			'xl/drawings/_rels/d.xml.rels',
			rels((i) => `../charts/c${i}.xml`),
		);
		for (const i of indices) zip.file(`xl/charts/c${i}.xml`, chartPart);
		const workbook = await loadXlsx(await zipBytes(zip));
		const charts = workbook.sheets[0]!.drawings.flatMap((drawing) =>
			drawing.kind === 'chart' ? [drawing] : [],
		);
		expectBudgeted(charts.map((chart) => chart.series[0]!.values.length));
		expect(charts.every((chart) => chart.editBlocked === undefined)).toBe(true);
	});

	it('is shared by the charts of a Word document', async () => {
		const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
		const drawing = (i: number) =>
			`<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="${i + 1}" name="c${i}"/><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rIdC${i}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:r="${R}"><w:body>${indices.map(drawing).join('')}<w:sectPr/></w:body></w:document>`,
		);
		zip.file(
			'word/_rels/document.xml.rels',
			rels((i) => `charts/c${i}.xml`),
		);
		for (const i of indices) zip.file(`word/charts/c${i}.xml`, chartPart);
		zip.file('[Content_Types].xml', `<Types xmlns="${CT}">${overrides('word/charts')}</Types>`);
		const { model } = await loadDocx(await zipBytes(zip));
		const charts = chartsIn(documentDrawingBlocks(model as DocumentModel));
		expectBudgeted(
			charts.map(
				(chart) =>
					chartCacheValues(chart.chartSpace!.plotArea.groups[0]!.series[0]!.values!.cache).length,
			),
		);
	});

	it('is shared by the charts of a presentation', async () => {
		const fixture = readFileSync(
			path.join(import.meta.dirname, '../pptx/__tests__/fixtures/embedded-assets-sample.pptx'),
		);
		const zip = await JSZip.loadAsync(fixture);
		const slide = Object.keys(zip.files).find((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))!;
		const slideRels = slide.replace('slides/', 'slides/_rels/') + '.rels';
		const frame = (i: number) =>
			`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${9000 + i}" name="c${i}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rIdC${i}"/></a:graphicData></a:graphic></p:graphicFrame>`;
		const edit = async (name: string, change: (xml: string) => string) =>
			zip.file(name, change(await zip.file(name)!.async('string')));
		await edit(slide, (xml) =>
			xml.replace('</p:spTree>', `${indices.map(frame).join('')}</p:spTree>`),
		);
		await edit(slideRels, (xml) =>
			xml.replace(
				'</Relationships>',
				indices
					.map(
						(i) => `<Relationship Id="rIdC${i}" Type="${R}/chart" Target="../charts/c${i}.xml"/>`,
					)
					.join('') + '</Relationships>',
			),
		);
		await edit('[Content_Types].xml', (xml) =>
			xml.replace('</Types>', `${overrides('ppt/charts')}</Types>`),
		);
		for (const i of indices) zip.file(`ppt/charts/c${i}.xml`, chartPart);
		const bytes = await zipBytes(zip);
		const data = await new PptxHandler().load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const charts = data.slides
			.flatMap((entry) => entry.elements)
			.filter((element): element is ChartPptxElement => element.type === 'chart');
		expectBudgeted(charts.map((chart) => chart.chartData!.series[0]!.values.length));
	});
});
