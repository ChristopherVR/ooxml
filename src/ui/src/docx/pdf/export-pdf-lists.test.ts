import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import { describe, expect, it } from 'vitest';
import { exportDocxPdf } from './export-pdf';
import { listNumbering } from './pdf-lists';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const measurer = createFakeMeasurer();

/** The file's objects by number, and its decoded page content streams. */
function inspect(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const objects = new Map<number, string>();
	for (const match of raw.matchAll(/(\d+) 0 obj\n([^]*?)\nendobj/g))
		objects.set(Number(match[1]), match[2]!);
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
	return { raw, objects, content: streams.join('\n') };
}

/** The structure tree as `Type(kids...)`, with `*` for marked content and `@` for annotations. */
function structure(pdf: ReturnType<typeof inspect>): string {
	const root = /\/Type \/StructTreeRoot \/K \[(\d+) 0 R\]/.exec(pdf.raw)!;
	const show = (id: number): string => {
		const body = pdf.objects.get(id)!;
		const numbering = /\/A << \/O \/List \/ListNumbering \/(\w+) >>/.exec(body)?.[1];
		const type = /\/S \/(\w+)/.exec(body)![1]! + (numbering ? `:${numbering}` : '');
		const kids = /\/K \[(.*)\] >>$/.exec(body)![1]!;
		const parts: string[] = [];
		for (const kid of kids.matchAll(/<< \/Type \/(MCR|OBJR)[^>]*>>|(\d+) 0 R/g))
			parts.push(kid[1] === 'MCR' ? '*' : kid[1] === 'OBJR' ? '@' : show(Number(kid[2])));
		return parts.length ? `${type}(${parts.join(' ')})` : type;
	};
	return show(Number(root[1]));
}

async function docx(body: string, files: Record<string, string> = {}, rels = '', refs = '') {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
	);
	for (const [name, data] of Object.entries(files)) zip.file(name, data);
	return (await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model;
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (text: string) => `<w:p>${run(text)}</w:p>`;
const item = (numId: number, level: number, text: string) =>
	`<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>${run(text)}</w:p>`;
const level = (ilvl: number, format: string, text: string) =>
	`<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
const numbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${level(0, 'decimal', '%1.')}${level(1, 'lowerLetter', '%2)')}</w:abstractNum><w:abstractNum w:abstractNumId="1">${level(0, 'bullet', '\uf0b7')}${level(1, 'bullet', 'o')}</w:abstractNum><w:abstractNum w:abstractNumId="2">${level(0, 'upperRoman', '%1.')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>`;
const numberingRel = `<Relationship Id="rN" Type="${R}/numbering" Target="numbering.xml"/>`;

describe('exportDocxPdf lists', () => {
	it('tags numbered and bulleted paragraphs as nested lists', async () => {
		const model = await docx(
			[
				para('Intro'),
				item(1, 0, 'One'),
				item(1, 1, 'One a'),
				item(1, 1, 'One b'),
				item(1, 0, 'Two'),
				item(3, 0, 'Roman'),
				para('Between'),
				item(2, 0, 'Dot'),
				item(2, 1, 'Ring'),
				`<w:tbl><w:tr><w:tc>${item(1, 0, 'Cell one')}${item(1, 0, 'Cell two')}</w:tc></w:tr></w:tbl>`,
			].join(''),
			{ 'word/numbering.xml': numbering },
			numberingRel,
		);
		const pdf = inspect((await exportDocxPdf(model, { measurer })).bytes);
		const li = 'LI(Lbl(*) LBody(*))';
		expect(structure(pdf)).toBe(
			[
				'Document(P(*)',
				`L:Decimal(LI(Lbl(*) LBody(* L:LowerAlpha(${li} ${li}))) ${li})`,
				`L:UpperRoman(${li})`,
				'P(*)',
				`L:Disc(LI(Lbl(*) LBody(* L:Circle(${li}))))`,
				`Table(TR(TD(L:Decimal(${li} ${li})))))`,
			].join(' '),
		);
		// The marker is its own marked content, before the item's text.
		expect(pdf.content).toMatch(
			/\/Lbl <<\/MCID (\d+)>> BDC\n(?:[^\n]*\n)*?BT [^\n]*\(1\.\) Tj[^]*?EMC\n\/LBody <<\/MCID \d+>> BDC\n(?:[^\n]*\n)*?BT [^\n]*\(One\) Tj/,
		);
		expect(pdf.content).toMatch(/\/Lbl <<\/MCID \d+>> BDC\n(?:[^\n]*\n)*?BT [^\n]*\(a\\\)\) Tj/);
	});

	it('keeps a list that runs onto the next page in one L', async () => {
		const items = Array.from({ length: 70 }, (_, index) => item(1, 0, `Item ${index}`));
		const model = await docx(items.join(''), { 'word/numbering.xml': numbering }, numberingRel);
		const pdf = inspect((await exportDocxPdf(model, { measurer })).bytes);
		expect(pdf.raw).toMatch(/\/Type \/Pages \/Kids \[\d+ 0 R \d+ 0 R\]/);
		const li = 'LI(Lbl(*) LBody(*))';
		expect(structure(pdf)).toBe(`Document(L:Decimal(${Array(70).fill(li).join(' ')}))`);
	});

	it('keeps numbered headings as headings', async () => {
		const model = await docx(
			`<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${run('Scope')}</w:p>`,
			{ 'word/numbering.xml': numbering },
			numberingRel,
		);
		const pdf = inspect((await exportDocxPdf(model, { measurer })).bytes);
		expect(structure(pdf)).toBe('Document(H1(*))');
	});

	it('maps Word number formats and bullets to ListNumbering', () => {
		expect(listNumbering('decimal', '%1.')).toBe('Decimal');
		expect(listNumbering('decimalZero', '%1.')).toBe('Decimal');
		expect(listNumbering('upperLetter', '%1.')).toBe('UpperAlpha');
		expect(listNumbering('lowerRoman', '%1.')).toBe('LowerRoman');
		expect(listNumbering('bullet', '\uf0b7')).toBe('Disc');
		expect(listNumbering('bullet', 'o')).toBe('Circle');
		expect(listNumbering('bullet', '\uf0a7')).toBe('Square');
		expect(listNumbering('bullet', '\u2713')).toBe('Disc');
	});
});

describe('exportDocxPdf links in headers and footers', () => {
	it('makes header and footer links clickable on every page', async () => {
		const header = `<w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:p>${run('Visit ')}<w:hyperlink r:id="rL">${run('our site')}</w:hyperlink></w:p></w:hdr>`;
		const footer = `<w:ftr xmlns:w="${W}"><w:p>${run('See ')}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF target \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run('Target')}<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run('1')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`;
		const model = await docx(
			`${para('First')}<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:bookmarkStart w:id="1" w:name="target"/>${run('Target')}<w:bookmarkEnd w:id="1"/></w:p>`,
			{
				'word/header1.xml': header,
				'word/footer1.xml': footer,
				'word/_rels/header1.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rL" Type="${R}/hyperlink" Target="https://example.com/h" TargetMode="External"/></Relationships>`,
			},
			`<Relationship Id="rH" Type="${R}/header" Target="header1.xml"/><Relationship Id="rF" Type="${R}/footer" Target="footer1.xml"/>`,
			'<w:headerReference w:type="default" r:id="rH"/><w:footerReference w:type="default" r:id="rF"/>',
		);
		const pdf = inspect((await exportDocxPdf(model, { measurer })).bytes);
		const pageIds = /\/Type \/Pages \/Kids \[(\d+) 0 R (\d+) 0 R\]/.exec(pdf.raw)!.slice(1, 3);
		for (const pageId of pageIds) {
			const page = pdf.objects.get(Number(pageId))!;
			const annots = [.../\/Annots \[([^\]]*)\]/.exec(page)![1]!.matchAll(/(\d+) 0 R/g)].map(
				(match) => pdf.objects.get(Number(match[1]))!,
			);
			expect(annots).toHaveLength(2);
			expect(annots[0]).toContain('/A << /S /URI /URI (https://example.com/h) >>');
			expect(annots[0]).toContain('/Contents (our site)');
			// The cross-reference goes to the bookmark on page 2.
			expect(annots[1]).toContain(`/Dest [${pageIds[1]} 0 R /XYZ`);
			expect(annots[1]).toContain('/Contents (Target)');
			expect(annots[1]).toMatch(/\/StructParent \d+/);
		}
		// Their text stays a pagination artifact; each annotation has its own Link element, after
		// the document's content.
		expect(structure(pdf)).toBe('Document(P(*) P(*) Link(@) Link(@) Link(@) Link(@))');
		expect(pdf.content).toMatch(
			/\/Artifact <<\/Type \/Pagination \/Subtype \/Header>> BDC\n[^]*?\(site\) Tj/,
		);
		expect(pdf.content).not.toMatch(/\/Link <<\/MCID/);
		// Without tags, the links are still there.
		const plain = inspect((await exportDocxPdf(model, { measurer, taggedPdf: false })).bytes);
		expect(plain.raw.match(/\/Subtype \/Link/g)).toHaveLength(4);
	});
});
