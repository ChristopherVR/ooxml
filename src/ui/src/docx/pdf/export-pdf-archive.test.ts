import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import { describe, expect, it } from 'vitest';
import { exportDocxPdf } from './export-pdf';
import { corpusFonts } from './__fixtures__/pdfa-corpus';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const measurer = createFakeMeasurer();
/** The synthetic font under one family only, so other families are substitutions. */
const fonts = corpusFonts().filter((font) => font.family === 'Arial' && !font.bold && !font.italic);

/** The raw file and its decoded page content streams. */
function inspect(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype|\/Length1|\/N 3/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return { raw, content: streams.filter((stream) => !stream.includes('begincmap')).join('\n') };
}

async function docx(body: string, header?: string) {
	const zip = new JSZip();
	const reference = header ? '<w:headerReference w:type="default" r:id="rH"/>' : '';
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr>${reference}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	if (header) {
		zip.file(
			'word/_rels/document.xml.rels',
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rH" Type="${R}/header" Target="header1.xml"/></Relationships>`,
		);
		zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}" xmlns:r="${R}">${header}</w:hdr>`);
	}
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const paragraph = (text: string, family: string) =>
	`<w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

describe('exportDocxPdf PDF/A', () => {
	it('embeds text in families with no font file through a supplied font, as a substitution', async () => {
		const { model } = await docx(paragraph('Segoe text', 'Segoe UI') + paragraph('Own', 'Arial'));
		const archived = await exportDocxPdf(model, { measurer, fonts, pdfA: true });
		const pdf = inspect(archived.bytes);
		expect(pdf.raw).not.toMatch(/\/Subtype \/Type1\b/);
		expect(pdf.raw).toContain('<pdfaid:conformance>U</pdfaid:conformance>');
		expect(archived.report.substitutedFonts).toEqual([
			{ family: 'Segoe UI', standardFont: 'Arial' },
		]);
		expect(archived.report.pdfA).toEqual({ conformant: true, level: 'U', issues: [] });
		// Without PDF/A the document's other families keep the standard fonts, as before.
		const plain = await exportDocxPdf(model, { measurer, fonts });
		expect(inspect(plain.bytes).raw).toContain('/BaseFont /Helvetica');
	});

	it('embeds chart text, whose family no font file names, for PDF/A', async () => {
		const file = join(import.meta.dirname, '../../../../core/docx/__fixtures__/chart.docx');
		const { model } = await loadDocx(new Uint8Array(readFileSync(file)));
		const archived = await exportDocxPdf(model, { measurer, fonts, pdfA: true });
		const pdf = inspect(archived.bytes);
		expect(pdf.raw).not.toMatch(/\/Subtype \/Type1\b/);
		expect(archived.report.pdfA?.conformant).toBe(true);
		expect(inspect((await exportDocxPdf(model, { measurer, fonts })).bytes).raw).toContain(
			'/BaseFont /Helvetica',
		);
	});

	it('names only the families written in standard fonts when the file does not conform', async () => {
		const { model } = await docx(paragraph('AB', 'Segoe UI') + paragraph('\u03b1\u03b2', 'Arial'));
		const { report } = await exportDocxPdf(model, { measurer, fonts, pdfA: true });
		// Greek is in no supplied font: Arial's text falls back to Helvetica; Segoe UI is embedded.
		expect(report.pdfA).toEqual({
			conformant: false,
			level: 'B',
			issues: [expect.stringMatching(/^Text in Arial uses standard PDF fonts/)],
		});
	});

	it('marks a watermark as an artifact in a tagged PDF', async () => {
		const watermark =
			'<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="PowerPlusWaterMarkObject357" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;width:412pt;height:165pt;rotation:315;z-index:-251658752" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/></v:shape></w:pict></w:r>';
		const { model } = await docx(paragraph('Body', 'Arial'), `<w:p>${watermark}</w:p>`);
		const tagged = inspect((await exportDocxPdf(model, { measurer, fonts, pdfA: true })).bytes);
		// PDF/UA 7.1: every content item is tagged or an artifact; the watermark was neither.
		expect(tagged.content).toMatch(
			/^[^\n]* cm\n\/Artifact <<\/Type \/Pagination>> BDC\nq\n\/GS\d gs\n[^]*?\] TJ ET[^\n]*\nQ\nEMC\n/,
		);
		const untagged = inspect(
			(await exportDocxPdf(model, { measurer, fonts, taggedPdf: false })).bytes,
		);
		expect(untagged.content).not.toContain('/Artifact');
	});
});
