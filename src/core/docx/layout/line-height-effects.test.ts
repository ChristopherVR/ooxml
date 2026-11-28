import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import type { TextMeasurer } from './measure';

// The reporter's .docx files for #394 and #397 are not attached to the issues, so the documents
// are generated here with the issues' properties. Word 16.0.20625 measurements, from the issues:
// a default-position superscript leaves the line pitch of 12 pt text unchanged (#394); a 14 pt
// run with `w14:shadow` lowers its baseline and the next line by the extra room of the shadow
// (#397: blurRad 50800 and dist 38100 at 45 degrees give 0.72 and 2.76 pt).
const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Arial-like metrics: ascent 0.905 em and descent 0.212 em.
const measurer: TextMeasurer = {
	widthOf: (text, font) => text.length * font.sizePx * 0.5,
	lineHeightOf: (font) => font.sizePx * 1.117,
	ascentOf: (font) => font.sizePx * 0.905,
};

async function baselines(body: string): Promise<number[]> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:w14="${w14}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const [page] = layoutDocumentModel(model, measurer).pages;
	return (page?.columns ?? []).flatMap((column) =>
		column.blocks.flatMap((block) =>
			block.kind === 'paragraph'
				? block.lines.map((line) => (block.yPx + line.yPx + (line.baselinePx ?? 0)) * 0.75)
				: [],
		),
	);
}

const fonts = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>';
const spacing =
	'<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>';
const line = (text: string, props: string, tail = '') =>
	`<w:p>${spacing}<w:r><w:rPr>${fonts}${props}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>${tail}</w:p>`;
const pitches = (values: number[]) => values.slice(1).map((value, i) => value - (values[i] ?? 0));

it('does not make a line taller for a default-position superscript (#394)', async () => {
	const sup = (props: string) =>
		Array.from(
			{ length: 5 },
			() =>
				`<w:p>${spacing}<w:r><w:rPr>${fonts}<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">from the 10</w:t></w:r><w:r><w:rPr>${fonts}<w:sz w:val="24"/>${props}</w:rPr><w:t>th</w:t></w:r><w:r><w:rPr>${fonts}<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve"> hour</w:t></w:r></w:p>`,
		).join('');
	const control = pitches(await baselines(sup('')));
	const superscript = pitches(await baselines(sup('<w:vertAlign w:val="superscript"/>')));
	expect(control).toHaveLength(4);
	for (const [i, value] of superscript.entries())
		expect(Math.abs(value - (control[i] ?? NaN))).toBeLessThan(0.01);
});

async function shadowSteps(blurRad: number, dist: number) {
	const effect = `<w14:shadow w14:blurRad="${blurRad}" w14:dist="${dist}" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:shadow>`;
	const body = (props: string) =>
		line('Title', `<w:b/><w:sz w:val="28"/>${props}`) + line('Next', '<w:sz w:val="22"/>');
	const [plainTitle = 0, plainNext = 0] = await baselines(body(''));
	const [title = 0, next = 0] = await baselines(body(effect));
	return { title: title - plainTitle, next: next - plainNext };
}

it('adds the room of a w14:shadow to the line height, as Word does (#397)', async () => {
	// Word: baseline +0.72 pt and next line +2.76 pt for blurRad 50800 and dist 38100.
	const both = await shadowSteps(50800, 38100);
	expect(Math.abs(both.title - 0.72)).toBeLessThan(0.15);
	expect(Math.abs(both.next - 2.76)).toBeLessThan(0.15);
	// Word: blurRad 101600 alone moves the baseline 2.64 pt and the next line 5.52 pt.
	const blur = await shadowSteps(101600, 0);
	expect(Math.abs(blur.title - 2.64)).toBeLessThan(0.15);
	expect(Math.abs(blur.next - 5.52)).toBeLessThan(0.15);
	// Word: dist 76200 alone leaves the baseline and moves the next line 1.44 pt.
	const offset = await shadowSteps(0, 76200);
	expect(Math.abs(offset.title)).toBeLessThan(0.15);
	expect(Math.abs(offset.next - 1.44)).toBeLessThan(0.15);
	expect(await shadowSteps(0, 0)).toEqual({ title: 0, next: 0 });
});
