// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { deflateSync, inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	anchor,
	CHART_PARTS,
	cellFloatCases,
	cellFloatPackage,
	type CellFloatCase,
} from '../__fixtures__/cell-floats/cases';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;
const FIXTURES = join(import.meta.dirname, '../__fixtures__');
const measurer = createFakeMeasurer();
/** A 1 x 1 opaque black PNG. */
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

/** The decoded page content of a PDF, its standard-font text and its image objects. */
function inspect(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	const images: Uint8Array[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		const data = bytes.subarray(start, start + length);
		const flate = /\/Filter \/FlateDecode/.test(match[1]!);
		if (/\/Subtype \/Image/.test(match[1]!)) {
			if (flate && /DeviceRGB/.test(match[1]!)) images.push(inflateSync(data));
		} else if (flate && !/\/Length1/.test(match[1]!))
			streams.push(inflateSync(data).toString('latin1'));
	}
	const content = streams.join('\n');
	const text = [...content.matchAll(/\(((?:\\.|[^\\)])*)\) Tj/g)]
		.map((m) => m[1]!.replace(/\\(.)/g, '$1'))
		.join('|');
	return { raw, content, text, images };
}

/** A one-section document with an optional default header (and its media). */
async function docx(
	body: string,
	options: { header?: string; sectPr?: string; png?: Uint8Array } = {},
) {
	const zip = new JSZip();
	const headerRef = options.header ? '<w:headerReference w:type="default" r:id="rH"/>' : '';
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${body}<w:sectPr>${headerRef}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>${options.sectPr ?? ''}</w:sectPr></w:body></w:document>`,
	);
	if (options.header) {
		zip.file(
			'[Content_Types].xml',
			'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
		);
		zip.file(
			'word/_rels/document.xml.rels',
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rH" Type="${R}/header" Target="header1.xml"/></Relationships>`,
		);
		zip.file('word/header1.xml', `<w:hdr ${NS}>${options.header}</w:hdr>`);
		zip.file(
			'word/_rels/header1.xml.rels',
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${R}/image" Target="media/image1.png"/></Relationships>`,
		);
		zip.file('word/media/image1.png', options.png ?? PNG);
	}
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

async function cellFloatDocx(entry: CellFloatCase) {
	const chart = await JSZip.loadAsync(readFileSync(join(FIXTURES, 'chart.docx')));
	const parts: Record<string, string> = {};
	for (const part of CHART_PARTS) parts[part] = await chart.file(part)!.async('string');
	const zip = new JSZip();
	for (const [path, xml] of Object.entries(cellFloatPackage(entry, parts))) zip.file(path, xml);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const omittedKinds = (report: { omitted: { kind: string }[] }) => report.omitted.map((o) => o.kind);

afterEach(() => vi.unstubAllGlobals());

describe('exportDocxPdf graphics', () => {
	it('paints the bordered paragraph below its top-and-bottom picture', async () => {
		const fixture = join(
			import.meta.dirname,
			'../layout/fixtures/float-in-bordered-paragraph/issue-352-picture-in-bordered-paragraph.docx',
		);
		const { model, media } = await loadDocx(readFileSync(fixture));
		const { bytes } = await exportDocxPdf(model, { measurer, ...(media ? { media } : {}) });
		const { content } = inspect(bytes);
		expect(content).toMatch(/\/Im\d+ Do/);
		const shading = /([\d.]+) ([\d.]+) 616\.8 ([\d.]+) re\nf/.exec(content);
		expect(Number(shading?.[2])).toBeCloseTo(296.47, 1);
		expect(Number(shading?.[3])).toBeCloseTo(56.95, 1);
	});

	it('writes a chart as vector paths with real text, never as a picture', async () => {
		const { model } = await loadDocx(new Uint8Array(readFileSync(join(FIXTURES, 'chart.docx'))));
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		expect(pdf.raw).not.toContain('/Subtype /Image');
		// Gridlines and bars are paths; the title and axis labels are text.
		expect(pdf.content).toMatch(/ m\n[\d. ]+ l\n[SB]\n/);
		expect(pdf.content).toMatch(/ l\nh\nf\n/);
		expect(pdf.text).toContain('Sales by region');
		expect(pdf.text).toContain('North');
		expect(omittedKinds(report)).not.toContain('chart');
	});

	it('writes the same chart without a DOMParser, through the DOM-free XML parser', async () => {
		const { model } = await loadDocx(new Uint8Array(readFileSync(join(FIXTURES, 'chart.docx'))));
		const withDom = inspect((await exportDocxPdf(model, { measurer })).bytes);
		vi.stubGlobal('DOMParser', undefined);
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		expect(omittedKinds(report)).not.toContain('chart');
		expect(inspect(bytes).content).toBe(withDom.content);
	});

	it('draws floating objects in table cells and a text box with its laid-out lines', async () => {
		const shapes = await cellFloatDocx(cellFloatCases()['textbox-cell']!);
		const { bytes, report } = await exportDocxPdf(shapes.model, { measurer });
		const pdf = inspect(bytes);
		// The red shape is a filled path; its text is real text, clipped to its text rectangle.
		expect(pdf.content).toMatch(/1 0 0 rg\n[^]*?\nf/);
		expect(pdf.text).toContain('alpha');
		expect(pdf.text).toContain('beta');
		expect(report.omitted).toEqual([]);
		const chart = await cellFloatDocx(cellFloatCases()['chart-cell']!);
		const charted = await exportDocxPdf(chart.model, { measurer });
		expect(inspect(charted.bytes).raw).not.toContain('/Subtype /Image');
		expect(charted.report.omitted).toEqual([]);
	});

	it('draws a header text watermark turned and half transparent', async () => {
		const watermark =
			'<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="PowerPlusWaterMarkObject357" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;width:412pt;height:165pt;rotation:315;z-index:-251658752" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/></v:shape></w:pict></w:r>';
		const { model } = await docx(para('Body'), {
			header: `<w:p>${watermark}<w:r><w:t>Head</w:t></w:r></w:p>`,
		});
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		expect(pdf.text.split('|')).toEqual(['DRAFT', 'Head', 'Body']);
		expect(pdf.content).toMatch(/\/GS\d gs\n0\.707107 -0\.707107 0\.707107 0\.707107 408 528 cm/);
		expect(pdf.raw).toMatch(/\/GS\d << \/ca 0\.5 \/CA 0\.5 >>/);
		expect(omittedKinds(report)).not.toContain('watermark');
	});

	it('draws a washed-out picture watermark centred on the margins', async () => {
		const picture = `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="WordPictureWatermark7" o:spid="_x0000_s2050" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:468pt;height:234pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f"><v:imagedata r:id="rId5" o:title="logo" gain="19661f" blacklevel="22938f"/></v:shape></w:pict></w:r>`;
		const loaded = await docx(para('Body'), { header: `<w:p>${picture}</w:p>` });
		const { bytes, report } = await exportDocxPdf(loaded.model, {
			measurer,
			media: loaded.media!,
		});
		const pdf = inspect(bytes);
		// 468 x 234 pt centred on the 6.5 x 9 in text area of a Letter page, 96 px per inch.
		expect(pdf.content).toContain('624 0 0 -312 96 684 cm\n/Im0 Do');
		// The black pixel is washed out to light grey (0.8), as Word does; white stays white.
		expect([...pdf.images[0]!]).toEqual([204, 204, 204]);
		expect(report.omitted).toEqual([]);
	});

	it('keeps white pixels white in a washed-out picture watermark', async () => {
		const picture = `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="WordPictureWatermark7" o:spid="_x0000_s2050" type="#_x0000_t75" style="position:absolute;width:100pt;height:50pt;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"><v:imagedata r:id="rId5" gain="19661f" blacklevel="22938f"/></v:shape></w:pict></w:r>`;
		// One black and one white pixel.
		const png = Buffer.from(
			'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGNgYGD4//8/AAYBAv4CsjmuAAAAAElFTkSuQmCC',
			'base64',
		);
		const loaded = await docx(para('Body'), { header: `<w:p>${picture}</w:p>`, png });
		const { bytes } = await exportDocxPdf(loaded.model, { measurer, media: loaded.media! });
		expect([...inspect(bytes).images[0]!]).toEqual([204, 204, 204, 255, 255, 255]);
	});

	it('strokes a connector with a zero height, level', async () => {
		const connector = `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>360000</wp:posOffset></wp:positionV><wp:extent cx="3600000" cy="0"/><wp:wrapNone/><wp:docPr id="1" name="Line 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3600000" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
		const { model } = await docx(`<w:p>${connector}<w:r><w:t>Below</w:t></w:r></w:p>`);
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		// 3600000 EMU is 378 px; red, 2 px wide, and level (no rise across its length).
		expect(pdf.content).toMatch(/1 0 0 RG\n2 w[^\n]*\n0 0 m\n378 0 l\nS/);
		expect(report.omitted).toEqual([]);
	});

	it('draws a floating shape anchored in the header', async () => {
		const shape = anchor(9, 'page:36', 'page:36', 'none', 1, { color: '0000FF' });
		const { model } = await docx(para('Body'), {
			header: `<w:p>${shape}<w:r><w:t>Head</w:t></w:r></w:p>`,
		});
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		// A 36 pt square, 36 pt from the page's top-left corner: 48 px at 48 px.
		expect(pdf.content).toMatch(/1 0 0 1 48 48 cm\n[^Q]*0 0 1 rg\n0 0 m\n48 0 l/);
		expect(report.omitted).toEqual([]);
	});

	it('draws page borders and line numbers', async () => {
		const border = 'w:val="single" w:sz="12" w:space="24" w:color="FF0000"';
		const { model } = await docx(para('one two three') + para('four'), {
			sectPr: `<w:pgBorders w:offsetFrom="page"><w:top ${border}/><w:left ${border}/><w:bottom ${border}/><w:right ${border}/></w:pgBorders><w:lnNumType w:countBy="1" w:restart="continuous"/>`,
		});
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		// 24 pt from each edge (32 px), 1.5 pt wide (2 px), red.
		expect(pdf.content).toContain('1 0 0 rg\n32 32 752 2 re\nf');
		expect(pdf.content).toContain('32 32 2 992 re\nf');
		expect(pdf.text).toBe('one |two |three|four|1|2');
		expect(report.omitted).toEqual([]);
		// Word's PDF: line numbers end 0.25 in left of the text, on the line's baseline.
		expect(pdf.content).toMatch(/1 0 0 -1 (\d+(\.\d+)?) [\d.]+ Tm[^\n]*\(1\) Tj/);
		const numberX = Number(/1 0 0 -1 ([\d.]+) [\d.]+ Tm[^\n]*\(1\) Tj/.exec(pdf.content)![1]);
		expect(numberX + measurer.widthOf('1', { family: 'Calibri', sizePx: 40 / 3 })).toBeCloseTo(
			72,
			3,
		);
	});

	it('measures text-relative page borders to the lines inner edge, as Word does', async () => {
		const border = 'w:val="single" w:sz="12" w:space="24" w:color="FF0000"';
		const { model } = await docx(para('x'), {
			sectPr: `<w:pgBorders w:offsetFrom="text"><w:top ${border}/><w:left ${border}/><w:bottom ${border}/><w:right ${border}/></w:pgBorders>`,
		});
		const pdf = inspect((await exportDocxPdf(model, { measurer })).bytes);
		// Word draws the 1.5 pt line from 46.5 to 48 pt: its inner edge 24 pt outside the 1 in margin.
		expect(pdf.content).toContain('62 62 692 2 re\nf');
	});
});
