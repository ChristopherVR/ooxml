/**
 * Test fixture: the documents `scripts/pdfa-validate.mjs` exports with `pdfA: true` and hands to
 * veraPDF, covering what the Word PDF export writes (embedded fonts, pictures with alpha and
 * JPEG, tables, links, bookmarks, the structure tree, headers and footers, footnotes, review
 * notes, charts, shapes and a translucent watermark). `pdfa-corpus.test.ts` builds them on every
 * run and checks the file-level PDF/A rules without veraPDF.
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { deflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import {
	anchor,
	CHART_PARTS,
	cellFloatCases,
	cellFloatPackage,
} from '../../../../../core/docx/__fixtures__/cell-floats/cases';
import type { PdfFontFile } from '../../../export/pdf-embedded-font';
import { syntheticTrueType } from '../../../pptx/export/__fixtures__/synthetic-truetype';
import { exportDocxPdf, type DocxPdfOptions, type DocxPdfResult } from '../export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;
const CORE = join(import.meta.dirname, '../../../../../core');
const FAMILIES = ['Calibri', 'Calibri Light', 'Cambria', 'Courier New', 'Arial', 'Times New Roman'];

/**
 * A synthetic font with a glyph for every printable ASCII character and a few Latin-1 ones, so
 * the corpus embeds all its text without a system font.
 */
export function corpusTrueType(): Uint8Array {
	const extra: [number, number][] = [];
	for (let code = 0x21; code <= 0x7e; code++)
		if (code !== 0x41 && code !== 0x42) extra.push([code, code % 2 ? 1 : 2]);
	for (const code of [0xe8, 0xe9, 0xfc, 0x2212]) extra.push([code, 2]);
	return syntheticTrueType(extra);
}

/** The font files to embed: `files` (paths) under every family the corpus uses, or the synthetic one. */
export function corpusFonts(files: readonly string[] = []): PdfFontFile[] {
	const data = files.length ? files.map((file) => new Uint8Array(readFileSync(file))) : [];
	const sources = data.length ? data : [corpusTrueType()];
	return FAMILIES.flatMap((family) =>
		sources.flatMap((bytes) =>
			[false, true].flatMap((bold) =>
				[false, true].map((italic) => ({ family, bold, italic, data: bytes })),
			),
		),
	);
}

/** A small RGBA PNG with a translucent column. */
function rgbaPng(width: number, height: number): Uint8Array {
	const table = Array.from({ length: 256 }, (_, n) => {
		let c = n;
		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		return c >>> 0;
	});
	const chunk = (type: string, data: Uint8Array) => {
		const out = new Uint8Array(12 + data.length);
		const view = new DataView(out.buffer);
		view.setUint32(0, data.length);
		out.set(Buffer.from(type, 'latin1'), 4);
		out.set(data, 8);
		let crc = 0xffffffff;
		for (const byte of out.subarray(4, 8 + data.length))
			crc = table[(crc ^ byte) & 255]! ^ (crc >>> 8);
		view.setUint32(8 + data.length, (crc ^ 0xffffffff) >>> 0);
		return out;
	};
	const header = new Uint8Array(13);
	new DataView(header.buffer).setUint32(0, width);
	new DataView(header.buffer).setUint32(4, height);
	header.set([8, 6, 0, 0, 0], 8);
	const rows: number[] = [];
	for (let y = 0; y < height; y++) {
		rows.push(0);
		for (let x = 0; x < width; x++) rows.push(30 * x, 200, 255 - 30 * y, x === 0 ? 96 : 255);
	}
	return Buffer.concat([
		Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk('IHDR', header),
		chunk('IDAT', new Uint8Array(deflateSync(new Uint8Array(rows)))),
		chunk('IEND', new Uint8Array()),
	]);
}

const run = (text: string, props = '') =>
	`<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (text: string, props = '') => `<w:p>${run(text, props)}</w:p>`;
const heading = (level: number, text: string) =>
	`<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>${run(text)}</w:p>`;
const inline = (rel: string, id: number, alt: string) =>
	`<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="635000"/><wp:docPr id="${id}" name="P${id}" descr="${alt}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="${rel}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

interface Parts {
	parts?: Record<string, string | Uint8Array>;
	rels?: string;
	sectPr?: string;
}

async function docx(body: string, { parts = {}, rels = '', sectPr = '' }: Parts = {}) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${body}<w:sectPr>${sectPr}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
	);
	for (const [name, data] of Object.entries(parts)) zip.file(name, data);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const textDocument = () =>
	docx(
		`${heading(1, 'Archive report')}${para('Plain text in the body, with an accent: caf\u00e9, and a minus: 3 \u2212 1.')}${para('Bold words', '<w:b/>')}${para('Italic words', '<w:i/>')}${heading(2, 'Second part')}${para('Numbers 0123456789 and (brackets) [too].')}`,
	);

async function picturesDocument() {
	const jpeg = new Uint8Array(
		readFileSync(join(CORE, 'docx/load/__fixtures__/legacy-doc/objects-input-b.jpg')),
	);
	return docx(
		`${para('A PNG with alpha and a JPEG:')}${inline('rP', 1, 'A gradient')}${inline('rJ', 2, 'A photo')}`,
		{
			parts: { 'word/media/a.png': rgbaPng(8, 6), 'word/media/b.jpeg': jpeg },
			rels: `<Relationship Id="rP" Type="${R}/image" Target="media/a.png"/><Relationship Id="rJ" Type="${R}/image" Target="media/b.jpeg"/>`,
		},
	);
}

function navigationDocument() {
	const cell = (text: string) => `<w:tc><w:p>${run(text)}</w:p></w:tc>`;
	const border = 'w:val="single" w:sz="4" w:color="000000"';
	const table = `<w:tbl><w:tblPr><w:tblBorders><w:top ${border}/><w:left ${border}/><w:bottom ${border}/><w:right ${border}/><w:insideH ${border}/><w:insideV ${border}/></w:tblBorders></w:tblPr><w:tr>${cell('Name')}${cell('Value')}</w:tr><w:tr>${cell('Alpha')}${cell('1')}</w:tr></w:tbl>`;
	const footnote = `<w:footnotes ${NS}><w:footnote w:id="1"><w:p>${run('A footnote.')}</w:p></w:footnote></w:footnotes>`;
	return docx(
		`${heading(1, 'Links')}<w:p>${run('See ')}<w:hyperlink r:id="rL">${run('the site')}</w:hyperlink>${run(' or ')}<w:hyperlink w:anchor="later">${run('later')}</w:hyperlink>${run('.')}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>${table}<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="1" w:name="later"/>${run('Later')}<w:bookmarkEnd w:id="1"/></w:p>${para('The end.')}`,
		{
			parts: {
				'word/header1.xml': `<w:hdr ${NS}>${para('Header text')}</w:hdr>`,
				'word/footer1.xml': `<w:ftr ${NS}>${para('Footer text')}</w:ftr>`,
				'word/footnotes.xml': footnote,
			},
			rels: `<Relationship Id="rL" Type="${R}/hyperlink" Target="https://example.com/a?b=1" TargetMode="External"/><Relationship Id="rH" Type="${R}/header" Target="header1.xml"/><Relationship Id="rF" Type="${R}/footer" Target="footer1.xml"/><Relationship Id="rN" Type="${R}/footnotes" Target="footnotes.xml"/>`,
			sectPr:
				'<w:headerReference w:type="default" r:id="rH"/><w:footerReference w:type="default" r:id="rF"/>',
		},
	);
}

const markupDocument = () =>
	docx(
		`<w:p>${run('Kept ')}<w:ins w:id="1" w:author="Ada"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="2" w:author="Bob"><w:r><w:delText>gone</w:delText></w:r></w:del><w:commentRangeStart w:id="7"/>${run(' noted')}<w:commentRangeEnd w:id="7"/></w:p>`,
		{
			parts: {
				'word/comments.xml': `<w:comments xmlns:w="${W}"><w:comment w:id="7" w:author="Cy" w:date="2026-03-04T05:06:07Z"><w:p><w:r><w:t>Check this</w:t></w:r></w:p></w:comment></w:comments>`,
			},
			rels: `<Relationship Id="rC" Type="${R}/comments" Target="comments.xml"/>`,
		},
	);

function watermarkDocument() {
	const watermark =
		'<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="PowerPlusWaterMarkObject357" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;width:412pt;height:165pt;rotation:315;z-index:-251658752" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/></v:shape></w:pict></w:r>';
	// A shape with a blurred shadow (a soft mask once the export draws the blur).
	const shape = anchor(9, 'page:36', 'page:36', 'none', 1, { color: '0000FF' }).replace(
		'<a:ln><a:noFill/></a:ln>',
		'<a:ln><a:noFill/></a:ln><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>',
	);
	// A gradient with a translucent stop: a shading under a soft mask, or a constant alpha.
	const gradient = anchor(10, 'page:120', 'page:36', 'none', 1, {
		color: '00AA00',
		widthPt: 72,
	}).replace(
		'<a:solidFill><a:srgbClr val="00AA00"/></a:solidFill>',
		'<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="40000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>',
	);
	return docx(para('Body under a watermark.'), {
		parts: { 'word/header1.xml': `<w:hdr ${NS}><w:p>${watermark}${shape}${gradient}${run('Head')}</w:p></w:hdr>` },
		rels: `<Relationship Id="rH" Type="${R}/header" Target="header1.xml"/>`,
		sectPr: '<w:headerReference w:type="default" r:id="rH"/>',
	});
}

async function chartDocument() {
	const chart = await JSZip.loadAsync(readFileSync(join(CORE, 'docx/__fixtures__/chart.docx')));
	const parts: Record<string, string> = {};
	for (const part of CHART_PARTS) parts[part] = await chart.file(part)!.async('string');
	const zip = new JSZip();
	const entry = cellFloatCases()['textbox-cell']!;
	for (const [path, xml] of Object.entries(cellFloatPackage(entry, parts))) zip.file(path, xml);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const chartFile = async () =>
	loadDocx(new Uint8Array(readFileSync(join(CORE, 'docx/__fixtures__/chart.docx'))));

export interface CorpusEntry {
	name: string;
	result: DocxPdfResult;
}

/** Exports every corpus document with `pdfA: true` and the given fonts (none: standard fonts). */
export async function buildPdfACorpus(fonts: readonly PdfFontFile[]): Promise<CorpusEntry[]> {
	const now = new Date(Date.UTC(2026, 0, 2, 3, 4, 5));
	const base: DocxPdfOptions = {
		measurer: createFakeMeasurer(),
		pdfA: true,
		now,
		language: 'en-US',
		fonts,
	};
	const documents: [string, Promise<Awaited<ReturnType<typeof loadDocx>>>, DocxPdfOptions][] = [
		['text', textDocument(), {}],
		['pictures', picturesDocument(), {}],
		['navigation', navigationDocument(), {}],
		['markup', markupDocument(), { markup: 'all' }],
		['watermark', watermarkDocument(), {}],
		['chart', chartFile(), {}],
		['shapes', chartDocument(), {}],
	];
	const out: CorpusEntry[] = [];
	for (const [name, loading, options] of documents) {
		const loaded = await loading;
		// Metadata the info dictionary and XMP must agree on, Unicode and XML specials included.
		loaded.model.properties =
			name === 'text'
				? {
						title: 'R\u00e9sum\u00e9 <draft> & "notes"\u0001',
						creator: 'Doe, Jane',
						subject: 'PDF/A \u2212 corpus',
						keywords: 'archive; pdf',
					}
				: { title: `Corpus ${name}`, creator: 'ooxml', subject: 'PDF/A' };
		const result = await exportDocxPdf(loaded.model, {
			...base,
			...(loaded.media ? { media: loaded.media } : {}),
			...options,
		});
		out.push({ name, result });
	}
	return out;
}
