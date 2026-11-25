import { deflateSync, inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import {
	describeCellFormatChange,
	describeRowFormatChange,
	describeSectionFormatChange,
	describeTableFormatChange,
	loadDocx,
	shapeSvg,
} from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import { describe, expect, it } from 'vitest';
import { PdfVectorPage } from '../../export/pdf-vector';
import { parseSvgMarkup } from '../../export/pdf-svg-dom';
import { drawSvgSubtree } from '../../export/pdf-svg';
import { exportDocxPdf } from './export-pdf';
import { corpusFonts } from './__fixtures__/pdfa-corpus';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:wps="${WPS}" xmlns:w14="${W14}"`;
const measurer = createFakeMeasurer();
const fonts = corpusFonts().filter((font) => font.family === 'Arial' && !font.bold && !font.italic);

const PNG = (() => {
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
	const header = new Uint8Array([0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0]);
	return Buffer.concat([
		Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk('IHDR', header),
		chunk('IDAT', new Uint8Array(deflateSync(new Uint8Array([0, 0, 0, 0])))),
		chunk('IEND', new Uint8Array()),
	]);
})();

function inspect(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (
			/\/Filter \/FlateDecode/.test(match[1]!) &&
			!/\/Subtype \/(Image|Form)|\/Length1|\/N 3/.test(match[1]!)
		)
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	const content = streams.filter((s) => !s.includes('begincmap')).join('\n');
	const text = [...content.matchAll(/\(((?:\\.|[^\\)])*)\) Tj/g)]
		.map((m) => m[1]!.replace(/\\(.)/g, '$1'))
		.join('|');
	return { raw, content, text };
}

async function makeDocx(
	body: string,
	extra: Record<string, string | Uint8Array> = {},
	rels = '',
	sectPr = '',
) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${body}<w:sectPr>${sectPr}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
	);
	for (const [path, data] of Object.entries(extra)) zip.file(path, data);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

describe('exportDocxPdf remaining Issue #318 gaps', () => {
	it('stretches text watermarks to their VML shape without an approximation note', async () => {
		const watermark =
			'<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="position:absolute;width:400pt;height:100pt;rotation:315" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="CONFIDENTIAL"/></v:shape></w:pict></w:r>';
		const { model } = await makeDocx(
			'<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
			{ 'word/header1.xml': `<w:hdr ${NS}><w:p>${watermark}</w:p></w:hdr>` },
			`<Relationship Id="rH" Type="${R}/header" Target="header1.xml"/>`,
			'<w:headerReference w:type="default" r:id="rH"/>',
		);
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		expect(pdf.text).toContain('CONFIDENTIAL');
		expect(report.notes.some((n) => n.includes('not stretched'))).toBe(false);
		// The rotation matrix is followed by a non-uniform scale matrix stretching text to 400x100pt.
		expect(pdf.content).toMatch(
			/0\.707107 -0\.707107 0\.707107 0\.707107 408 528 cm\n[\d.]+ 0 0 [\d.]+ 0 0 cm/,
		);
	});

	it('reads and paints Word shape shadows and picture shadows', async () => {
		const shapeDrawing = `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="S1" descr="Box"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
		const picDrawing = `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="P1" descr="Photo"/><a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="P1"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rImg"/></pic:blipFill><pic:spPr><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
		const loaded = await makeDocx(
			`<w:p>${shapeDrawing}${picDrawing}</w:p>`,
			{ 'word/media/img1.png': PNG },
			`<Relationship Id="rImg" Type="${R}/image" Target="media/img1.png"/>`,
		);
		const p = loaded.model.blocks[0]!;
		expect(p.type).toBe('paragraph');
		if (p.type === 'paragraph') {
			expect(p.runs[0]?.image?.shape?.effects?.effects[0]?.kind).toBe('outerShadow');
			expect(shapeSvg(p.runs[0]!.image!.shape!, 96, 48)).toContain('<feDropShadow');
			expect(p.runs[1]?.image?.effects?.effects[0]?.kind).toBe('outerShadow');
		}
		const { bytes, report } = await exportDocxPdf(loaded.model, {
			measurer,
			media: loaded.media!,
		});
		const pdf = inspect(bytes);
		expect(report.omitted).toEqual([]);
		// Soft-masked blurred shadow images are written for both the shape and the picture.
		expect(pdf.raw.match(/\/SMask \d+ 0 R/g)?.length).toBeGreaterThanOrEqual(2);
	});

	it('paints chart SVG text drop shadows instead of dropping them', async () => {
		const svgMarkup = `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="60" viewBox="0 0 200 60"><defs><filter id="ts"><feDropShadow dx="2" dy="2" stdDeviation="1.5" flood-color="#000000" flood-opacity="0.5"/></filter></defs><text x="10" y="30" font-family="Arial" font-size="16" fill="#222222" filter="url(#ts)">Chart Title</text></svg>`;
		const svg = (await parseSvgMarkup(svgMarkup))!;
		const page = new PdfVectorPage(200, 60);
		const notes = drawSvgSubtree(page, svg, new Map());
		expect(notes.includes('filterDropped')).toBe(false);
		expect(notes.includes('shadowBlurDropped')).toBe(false);
		expect(page.content).toContain('/ActualText ()');
		expect(page.content).toContain('(Chart Title) Tj');
	});

	it('reads and paints WordArt glow and reflection without duplicating extracted text', async () => {
		const wordArt = `<w:r><w:drawing><wp:inline><wp:extent cx="2743200" cy="914400"/><wp:docPr id="1" name="WA1" descr="Banner"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"/><a:noFill/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:color w:val="1769A5"/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="1769A5"/></w14:solidFill></w14:textFill><w14:glow w14:rad="63500"><w14:srgbClr w14:val="FFCC00"><w14:alpha w14:val="50000"/></w14:srgbClr></w14:glow><w14:reflection w14:blurRad="6350" w14:stA="50000" w14:endA="0" w14:dist="12700" w14:sy="-100000"/></w:rPr><w:t>GLOW</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr fromWordArt="1"><a:prstTxWarp prst="textArchUp"/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
		const { model } = await makeDocx(`<w:p>${wordArt}</w:p>`);
		const p = model.blocks[0]!;
		if (p.type === 'paragraph') {
			expect(p.runs[0]?.image?.textBoxContent?.wordArt?.glow).toMatchObject({
				color: 'FFCC00',
				radiusEmu: 63500,
			});
			expect(p.runs[0]?.image?.textBoxContent?.wordArt?.reflection).toMatchObject({
				startOpacity: 50000,
			});
		}
		const { bytes } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		expect(pdf.content).toContain('/ActualText ()');
		expect(pdf.content).toContain('(G) Tj');
	});

	it('tags lists in footnotes and text boxes and makes links in floating header text boxes clickable', async () => {
		const level = `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`;
		const numbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${level}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
		const listP = (text: string) =>
			`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
		const bodyBox = `<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="TB1" descr="Box"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></wps:spPr><wps:txbx><w:txbxContent>${listP('Box item')}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
		const headerFloatBox = `<w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" relativeHeight="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="457200"/><wp:wrapNone/><wp:docPr id="2" name="HBox"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"/><a:noFill/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:hyperlink r:id="rHL"><w:r><w:t>Header Link</w:t></w:r></w:hyperlink></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
		const { model } = await makeDocx(
			`<w:p><w:r><w:t>Main</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:p>${bodyBox}</w:p>`,
			{
				'word/numbering.xml': numbering,
				'word/footnotes.xml': `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${listP('Note item')}</w:footnote></w:footnotes>`,
				'word/header1.xml': `<w:hdr ${NS}><w:p>${headerFloatBox}</w:p></w:hdr>`,
				'word/_rels/header1.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rHL" Type="${R}/hyperlink" Target="https://example.com/hdr-box" TargetMode="External"/></Relationships>`,
			},
			`<Relationship Id="rN" Type="${R}/numbering" Target="numbering.xml"/><Relationship Id="rFn" Type="${R}/footnotes" Target="footnotes.xml"/><Relationship Id="rH" Type="${R}/header" Target="header1.xml"/>`,
			'<w:headerReference w:type="default" r:id="rH"/>',
		);
		const pdf = inspect((await exportDocxPdf(model, { measurer })).bytes);
		// Both the footnote list and the text box list are tagged as L > LI > Lbl + LBody.
		expect(pdf.raw.match(/\/S \/L /g)?.length).toBeGreaterThanOrEqual(2);
		// The hyperlink inside the floating header text box is clickable!
		expect(pdf.raw).toContain('/A << /S /URI /URI (https://example.com/hdr-box) >>');
	});

	it('emits balloons for table, row, cell and section property changes', async () => {
		const table = `<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:jc w:val="center"/><w:tblPrChange w:id="1" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr></w:tblPrChange></w:tblPr><w:tr><w:trPr><w:cantSplit/><w:trPrChange w:id="2" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:trPr/></w:trPrChange></w:trPr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="FFFF00"/><w:tcPrChange w:id="3" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:tcPr/></w:tcPrChange></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
		const sectChange =
			'<w:sectPrChange w:id="4" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:sectPrChange>';
		const { model } = await makeDocx(table, {}, '', sectChange);
		const tbl = model.blocks[0]!;
		expect(tbl.type).toBe('table');
		if (tbl.type === 'table') {
			expect(describeTableFormatChange(tbl)).toContain('Table alignment: Center');
			expect(describeRowFormatChange(tbl.rowProperties![0]!)).toContain('Prevent row from splitting');
			expect(describeCellFormatChange(tbl.rows[0]![0]!)).toContain('Shading: Yellow');
		}
		expect(describeSectionFormatChange(model.sections![0]!)).toContain('Margins: Top: 1"');
		const { bytes } = await exportDocxPdf(model, { measurer, markup: 'all' });
		const pdf = inspect(bytes);
		expect(pdf.text).toContain('Table');
		expect(pdf.text).toContain('Prevent');
		expect(pdf.text).toContain('Yellow');
		expect(pdf.text).toContain('Margins:');
	});

	it('supports PDF/A-2a and PDF/UA-1 and reports unmet requirements honestly', async () => {
		const coreXml =
			'<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Accessible Report</dc:title></cp:coreProperties>';
		const rootRels =
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>';
		const { model } = await makeDocx('<w:p><w:r><w:t>Hello</w:t></w:r></w:p>', {
			'_rels/.rels': rootRels,
			'docProps/core.xml': coreXml,
		});
		const full = await exportDocxPdf(model, {
			measurer,
			fonts,
			pdfA: true,
			pdfALevel: 'A',
			pdfUa: true,
			language: 'en-US',
		});
		const pdf = inspect(full.bytes);
		expect(full.report.pdfA).toEqual({ conformant: true, level: 'A', issues: [] });
		expect(full.report.pdfUa).toEqual({ conformant: true, issues: [] });
		expect(pdf.raw).toContain('<pdfaid:conformance>A</pdfaid:conformance>');
		expect(pdf.raw).toContain('<pdfuaid:part>1</pdfuaid:part>');
		expect(pdf.raw).toContain('/DisplayDocTitle true');

		// Missing language: neither 2a nor PDF/UA-1 is claimed; falls back to 2u.
		const missingLang = await exportDocxPdf(model, {
			measurer,
			fonts,
			pdfA: true,
			pdfALevel: 'A',
			pdfUa: true,
		});
		const missingRaw = inspect(missingLang.bytes).raw;
		expect(missingLang.report.pdfA?.conformant).toBe(false);
		expect(missingLang.report.pdfA?.level).toBe('U');
		expect(missingLang.report.pdfUa?.conformant).toBe(false);
		expect(missingRaw).toContain('<pdfaid:conformance>U</pdfaid:conformance>');
		expect(missingRaw).not.toContain('<pdfuaid:part>1</pdfuaid:part>');
	});
});
