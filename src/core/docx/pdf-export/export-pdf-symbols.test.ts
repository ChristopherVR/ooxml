// Symbol characters (`w:sym`) and symbol fonts in the Word PDF export (issue #365).
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { syntheticTrueType } from '../../text/truetype/test-support/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = createFakeMeasurer();

/** The raw file, its decoded content streams and the standard-font text they show. */
function inspectPdf(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype \/Image|\/Length1/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	const content = streams.join('\n');
	const text = [...content.matchAll(/\(((?:\\.|[^\\)])*)\) Tj/g)]
		.map((match) => match[1]!.replace(/\\(.)/g, '$1'))
		.join('');
	return { raw, content, text };
}

async function docxModel(body: string, extra: Record<string, string> = {}) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	for (const [name, data] of Object.entries(extra)) zip.file(name, data);
	return (await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model;
}

const sym = (font: string, char: string) => `<w:r><w:sym w:font="${font}" w:char="${char}"/></w:r>`;
const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;

describe('exportDocxPdf symbols', () => {
	it('writes a code below U+0100 in a text font as that character', async () => {
		// Word 16 shows Calibri 00A7 and 00B6 as the section and pilcrow signs.
		const model = await docxModel(
			`<w:p>${text('[')}${sym('Calibri', '00A7')}${text('] [')}${sym('Calibri', '00B6')}${text(']')}</w:p>`,
		);
		const { text: written } = inspectPdf((await exportDocxPdf(model, { measurer })).bytes);
		expect(written).toContain('[\xa7]');
		expect(written).toContain('[\xb6]');
		expect(written).not.toContain('?');
	});

	it('writes Symbol characters in the standard Symbol font without the font file', async () => {
		const model = await docxModel(`<w:p>${sym('Symbol', 'F061')}${sym('Symbol', 'F0B3')}</w:p>`);
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspectPdf(bytes);
		// Symbol's codes are the standard font's: a is alpha, 0xB3 is greater-or-equal.
		expect(pdf.raw).toMatch(/\/BaseFont \/Symbol >>/);
		expect(pdf.raw).not.toMatch(/\/BaseFont \/Symbol \/Encoding/);
		expect(pdf.text).toContain('a\xb3');
		expect(report.missingCharacters).toEqual([]);
	});

	it('embeds a supplied symbol font whose only character map is the symbol one', async () => {
		// A Symbol-like font: (3,0) cmap with U+F041 and U+F042 as its pictures.
		const font = syntheticTrueType(
			[
				[0xf041, 1],
				[0xf042, 2],
			],
			0,
		);
		const model = await docxModel(
			`<w:p>${sym('Wingdings', 'F041')}${sym('Wingdings', 'F042')}</w:p>`,
		);
		const { bytes, report } = await exportDocxPdf(model, {
			measurer,
			fonts: [{ family: 'Wingdings', data: font }],
		});
		const pdf = inspectPdf(bytes);
		expect(pdf.raw).toContain('/FontFile2');
		expect(report.embeddedFonts).toContain('Wingdings');
		// The ToUnicode map keeps the private-use codes, as Word's own export does.
		expect(pdf.content).toMatch(/<0001> <F041>/i);
	});
});

const bulletLevel = (id: number, char: string, font: string) =>
	`<w:abstractNum w:abstractNumId="${id}"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${char}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="${id + 1}"><w:abstractNumId w:val="${id}"/></w:num>`;
const bulletDocx = (levels: [string, string][]) =>
	docxModel(
		levels
			.map(
				(_, index) =>
					`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${index + 1}"/></w:numPr></w:pPr>${text('Item')}</w:p>`,
			)
			.join(''),
		{
			'word/numbering.xml': `<w:numbering xmlns:w="${W}">${levels.map(([char, font], index) => bulletLevel(index, char, font)).join('')}</w:numbering>`,
			'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
		},
	);

describe('exportDocxPdf symbol-font bullets (issue #366)', () => {
	it('writes Symbol bullets as their own pictures, not all as a bullet', async () => {
		// Word 16: Symbol F02D, F0B7 and F0AE are a minus sign, a bullet and a right arrow.
		const model = await bulletDocx([
			['', 'Symbol'],
			['', 'Symbol'],
			['', 'Symbol'],
		]);
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspectPdf(bytes);
		// Without the font file, in the standard Symbol font's codes for those pictures.
		expect(pdf.text).toContain('-');
		expect(pdf.text).toContain('\xb7');
		expect(pdf.text).toContain('\xae');
		expect(report.missingCharacters).toEqual([]);
	});

	it("draws a Wingdings check mark bullet with the supplied font's picture", async () => {
		const model = await bulletDocx([['', 'Wingdings']]);
		const font = syntheticTrueType([[0xf0fc, 1]], 0);
		const { bytes, report } = await exportDocxPdf(model, {
			measurer,
			fonts: [{ family: 'Wingdings', data: font }],
		});
		const pdf = inspectPdf(bytes);
		expect(report.embeddedFonts).toContain('Wingdings');
		// The glyph is Wingdings' F0FC; it extracts as U+2713, as in Word's PDF.
		expect(pdf.content).toMatch(/<0001> <2713>/i);
		expect(report.missingCharacters).toEqual([]);
	});
});
