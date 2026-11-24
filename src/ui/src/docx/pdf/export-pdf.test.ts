import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { deflateSync, inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import { describe, expect, it } from 'vitest';
import { syntheticTrueType } from '../../pptx/export/__fixtures__/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const measurer = createFakeMeasurer();

/** The decoded content streams of a PDF, one per page, and its raw text. */
function inspect(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		const data = bytes.subarray(start, start + length);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype \/Image|\/Length1/.test(match[1]!))
			streams.push(inflateSync(data).toString('latin1'));
	}
	const pages = (raw.match(/\/Type \/Page\b(?!s)/g) ?? []).length;
	// The standard-font text of each page, in drawing order.
	const texts = streams.map((stream) =>
		[...stream.matchAll(/\(((?:\\.|[^\\)])*)\) Tj/g)]
			.map((match) => match[1]!.replace(/\\(.)/g, '$1'))
			.join(''),
	);
	return { raw, streams, content: streams.join('\n'), pages, text: texts.join('\n'), texts };
}

async function docx(
	body: string,
	extra: Record<string, string | Uint8Array> = {},
	rels = '',
	references = '',
) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body}<w:sectPr>${references}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	if (rels)
		zip.file(
			'word/_rels/document.xml.rels',
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
		);
	for (const [name, data] of Object.entries(extra)) zip.file(name, data);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const paragraph = (text: string, props = '') =>
	`<w:p><w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

describe('exportDocxPdf', () => {
	it('writes real text with no images for a text-only document', async () => {
		const { model } = await docx(paragraph('Hello PDF') + paragraph('Second (line)', '<w:b/>'));
		const { bytes, blob, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		expect(pdf.raw.startsWith('%PDF-1.4')).toBe(true);
		expect(blob.type).toBe('application/pdf');
		expect(pdf.pages).toBe(1);
		expect(pdf.raw).toContain('/MediaBox [0 0 612 792]');
		expect(pdf.raw).not.toContain('/Subtype /Image');
		expect(pdf.text).toBe('Hello PDFSecond (line)');
		expect(pdf.raw).toContain('/BaseFont /Helvetica-Bold');
		expect(report.pages).toBe(1);
		expect(report.substitutedFonts).toContainEqual({
			family: 'Calibri',
			standardFont: 'Helvetica',
		});
	});

	it('embeds a supplied font as a subset with a ToUnicode map', async () => {
		const { model } = await docx(paragraph('AB BA', '<w:rFonts w:ascii="Synth" w:hAnsi="Synth"/>'));
		const { bytes, report } = await exportDocxPdf(model, {
			measurer,
			fonts: [{ family: 'Synth', data: syntheticTrueType() }],
		});
		const pdf = inspect(bytes);
		expect(pdf.raw).toContain('/Subtype /Type0');
		expect(pdf.raw).toContain('/FontFile2');
		expect(pdf.raw).toContain('/ToUnicode');
		// AB  then BA: the glyphs get CIDs in first use, the space stretched by kerning.
		expect(pdf.content).toMatch(/\[<00010002> <0003> -?[\d.]+\] TJ/);
		expect(pdf.content).toMatch(/\[<00020001>\] TJ/);
		expect(report.embeddedFonts).toEqual(['Synth']);
		expect(report.substitutedFonts).toEqual([]);
	});

	it('draws table borders and header and footer tables as vector paths (t10)', async () => {
		const file = join(
			import.meta.dirname,
			'../../../../core/docx/layout/fixtures/table-reports/t10-tableau-entete-pied.docx',
		);
		const { model } = await loadDocx(new Uint8Array(await readFile(file)));
		const { bytes } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		expect(pdf.raw).not.toContain('/Subtype /Image');
		for (const text of ['Corps de la lettre.', 'Cabinet Exemple', 'Toque P 000', 'Contrat'])
			expect(pdf.text).toContain(text);
		// Header text is WinAnsi-encoded: "référence" keeps its accented letters as real text.
		expect(pdf.text).toContain('Notre r\xe9f\xe9rence : DOS-2026-001');
		// Three tables of 4, 4 and 6 cells: each cell draws at least its top and left border.
		expect((pdf.content.match(/ re\nf/g) ?? []).length).toBeGreaterThanOrEqual(28);
	});

	it('fills page fields per page and honours the page range', async () => {
		const breakRun = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
		const footer = `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`;
		const body = `${paragraph('One')}${breakRun}${paragraph('Two')}${breakRun}${paragraph('Three')}`;
		const { model } = await docx(
			body,
			{ 'word/footer1.xml': footer },
			`<Relationship Id="rF" Type="${R}/footer" Target="footer1.xml"/>`,
			'<w:footerReference w:type="default" r:id="rF"/>',
		);
		const all = inspect((await exportDocxPdf(model, { measurer })).bytes);
		expect(all.pages).toBe(3);
		const some = await exportDocxPdf(model, { measurer, pages: { from: 2, to: 3 } });
		const pdf = inspect(some.bytes);
		expect(pdf.pages).toBe(2);
		expect(some.report.pages).toBe(2);
		expect(pdf.texts).toEqual(['TwoPage 2', 'ThreePage 3']);
		await expect(exportDocxPdf(model, { measurer, pages: { from: 9, to: 9 } })).rejects.toThrow(
			RangeError,
		);
	});

	it('keeps a PNG picture at its own resolution with its alpha, and a JPEG as it is', async () => {
		const png = rgbaPng(3, 2);
		const jpeg = new Uint8Array([
			0xff, 0xd8, 0xff, 0xc0, 0, 11, 8, 0, 2, 0, 3, 1, 1, 0x11, 0, 0xff, 0xda, 0, 8, 1, 1, 0, 0,
			0x3f, 0, 0x12, 0xff, 0xd9,
		]);
		const drawing = (rel: string, id: number) =>
			`<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="635000"/><wp:docPr id="${id}" name="P${id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="${rel}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
		const loaded = await docx(
			drawing('rP', 1) + drawing('rJ', 2),
			{ 'word/media/a.png': png, 'word/media/b.jpeg': jpeg },
			`<Relationship Id="rP" Type="${R}/image" Target="media/a.png"/><Relationship Id="rJ" Type="${R}/image" Target="media/b.jpeg"/>`,
		);
		const { bytes, report } = await exportDocxPdf(loaded.model, { measurer, media: loaded.media! });
		const pdf = inspect(bytes);
		expect(pdf.raw).toMatch(
			/\/Subtype \/Image \/Width 3 \/Height 2 \/ColorSpace \/DeviceRGB[^>]*\/SMask/,
		);
		expect(pdf.raw).toMatch(
			/\/Width 3 \/Height 2 \/ColorSpace \/DeviceGray \/BitsPerComponent 8 \/Filter \/DCTDecode/,
		);
		expect(pdf.content.match(/\/Im\d Do/g)).toHaveLength(2);
		expect(report.omitted).toEqual([]);
	});

	it('reports a picture it cannot read instead of dropping it silently', async () => {
		const loaded = await docx(
			`<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="635000"/><wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rX"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
			{ 'word/media/x.gif': new Uint8Array([71, 73, 70, 56, 57, 97, 1, 0, 1, 0, 0, 0, 0]) },
			`<Relationship Id="rX" Type="${R}/image" Target="media/x.gif"/>`,
		);
		const { report } = await exportDocxPdf(loaded.model, { measurer, media: loaded.media! });
		expect(report.omitted).toEqual([expect.objectContaining({ kind: 'picture', pages: [1] })]);
	});

	it('writes the document properties as Unicode text', async () => {
		const { model } = await docx(paragraph('x'));
		model.properties = { title: 'Résumé', creator: 'Ada' };
		const pdf = inspect((await exportDocxPdf(model, { measurer, language: 'fr-FR' })).bytes);
		expect(pdf.raw).toContain('/Title <FEFF005200E900730075006D00E9>');
		expect(pdf.raw).toContain('/Author (Ada)');
		expect(pdf.raw).toContain('/Lang (fr-FR)');
		const bare = inspect(
			(await exportDocxPdf(model, { measurer, documentProperties: false })).bytes,
		);
		expect(bare.raw).not.toContain('/Title');
	});

	it('stops when its signal is aborted', async () => {
		const { model } = await docx(paragraph('x'));
		const controller = new AbortController();
		controller.abort();
		await expect(exportDocxPdf(model, { measurer, signal: controller.signal })).rejects.toThrow();
	});
});

/** A small RGBA PNG with one translucent pixel. */
function rgbaPng(width: number, height: number): Uint8Array {
	const crcTable = Array.from({ length: 256 }, (_, n) => {
		let c = n;
		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		return c >>> 0;
	});
	const crc = (bytes: Uint8Array) => {
		let c = 0xffffffff;
		for (const b of bytes) c = crcTable[(c ^ b) & 255]! ^ (c >>> 8);
		return (c ^ 0xffffffff) >>> 0;
	};
	const chunk = (type: string, data: Uint8Array) => {
		const out = new Uint8Array(12 + data.length);
		const view = new DataView(out.buffer);
		view.setUint32(0, data.length);
		out.set(
			[...type].map((c) => c.charCodeAt(0)),
			4,
		);
		out.set(data, 8);
		view.setUint32(8 + data.length, crc(out.subarray(4, 8 + data.length)));
		return out;
	};
	const header = new Uint8Array(13);
	new DataView(header.buffer).setUint32(0, width);
	new DataView(header.buffer).setUint32(4, height);
	header.set([8, 6, 0, 0, 0], 8);
	const rows: number[] = [];
	for (let y = 0; y < height; y++) {
		rows.push(0);
		for (let x = 0; x < width; x++) rows.push(200, 10, 10, x === 0 && y === 0 ? 128 : 255);
	}
	const parts = [
		new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk('IHDR', header),
		chunk('IDAT', new Uint8Array(deflateSync(new Uint8Array(rows)))),
		chunk('IEND', new Uint8Array()),
	];
	return Uint8Array.from(parts.flatMap((part) => [...part]));
}
