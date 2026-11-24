import { writeFile } from 'node:fs/promises';
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import { describe, expect, it } from 'vitest';
import { syntheticTrueType } from '../../pptx/export/__fixtures__/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';
import { fieldLinkTarget } from './pdf-navigation';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const measurer = createFakeMeasurer();

/** The raw file and its decoded page content streams. */
function inspect(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		const data = bytes.subarray(start, start + length);
		if (
			/\/Filter \/FlateDecode/.test(match[1]!) &&
			!/\/Subtype \/(Image|Form)|\/Length1|\/N 3/.test(match[1]!)
		)
			streams.push(inflateSync(data).toString('latin1'));
	}
	return { raw, content: streams.join('\n') };
}

async function docx(body: string, extra: Record<string, string> = {}, rels = '', references = '') {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr>${references}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
	);
	for (const [name, data] of Object.entries(extra)) zip.file(name, data);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const heading = (level: number, text: string, bookmark?: string) =>
	`<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>${
		bookmark ? `<w:bookmarkStart w:id="${level}${text.length}" w:name="${bookmark}"/>` : ''
	}${run(text)}${bookmark ? `<w:bookmarkEnd w:id="${level}${text.length}"/>` : ''}</w:p>`;
const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

const linked = () =>
	docx(
		`${heading(1, 'Intro')}<w:p>${run('See ')}<w:hyperlink r:id="rL">${run('the site')}</w:hyperlink>${run(' and ')}<w:hyperlink w:anchor="target">${run('the target')}</w:hyperlink>${run(' or ')}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF target \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run('Details')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>${heading(2, 'Scope')}${pageBreak}${heading(1, 'Details', 'target')}${heading(2, 'More', 'Notes')}`,
		{},
		`<Relationship Id="rL" Type="${R}/hyperlink" Target="https://example.com/a?b=1" TargetMode="External"/>`,
	);

describe('exportDocxPdf navigation and structure', () => {
	it('turns hyperlinks and cross-references into link annotations', async () => {
		const { model } = await linked();
		const { bytes } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		await dump('links', bytes);
		expect(pdf.raw).toContain('/Subtype /Link');
		expect(pdf.raw).toContain('/A << /S /URI /URI (https://example.com/a?b=1) >>');
		// The anchor and the REF field go to page 2, where the bookmark is.
		const dests = [...pdf.raw.matchAll(/\/Subtype \/Link [^\n]*?\/Dest \[(\d+) 0 R \/XYZ/g)];
		expect(dests).toHaveLength(2);
		const kids = /\/Type \/Pages \/Kids \[(\d+) 0 R (\d+) 0 R\]/.exec(pdf.raw)!;
		for (const dest of dests) expect(dest[1]).toBe(kids[2]);
		expect(pdf.raw).toMatch(/\/Annots \[\d+ 0 R \d+ 0 R \d+ 0 R\]/);
	});

	it('builds the outline from headings, nested by level, or from Word bookmarks', async () => {
		const { model } = await linked();
		const pdf = inspect((await exportDocxPdf(model, { measurer })).bytes);
		expect(pdf.raw).toMatch(/\/Type \/Outlines \/First \d+ 0 R \/Last \d+ 0 R \/Count 4/);
		expect(pdf.raw).toMatch(/\/Title \(Intro\) [^\n]*\/First \d+ 0 R \/Last \d+ 0 R \/Count 1/);
		expect(pdf.raw).toContain('/Title (Scope)');
		expect(pdf.raw).toContain('/PageMode /UseOutlines');
		const marks = inspect(
			(await exportDocxPdf(model, { measurer, bookmarks: 'wordBookmarks' })).bytes,
		);
		expect([...marks.raw.matchAll(/\/Title \(([^)]*)\)/g)].map((m) => m[1])).toEqual([
			'target',
			'Notes',
		]);
		const none = inspect((await exportDocxPdf(model, { measurer, bookmarks: 'none' })).bytes);
		expect(none.raw).not.toContain('/Outlines');
	});

	it('tags headings, paragraphs, links and tables with marked content', async () => {
		const { model } = await linked();
		const table = await docx(
			`<w:tbl><w:tr><w:tc><w:p>${run('A1')}</w:p></w:tc><w:tc><w:p>${run('B1')}</w:p></w:tc></w:tr></w:tbl>`,
		);
		const pdf = inspect((await exportDocxPdf(model, { measurer })).bytes);
		expect(pdf.raw).toContain('/MarkInfo << /Marked true >>');
		expect(pdf.raw).toMatch(/\/StructTreeRoot \d+ 0 R/);
		expect(pdf.raw).toMatch(/\/Type \/StructTreeRoot \/K \[\d+ 0 R\] \/ParentTree \d+ 0 R/);
		for (const type of ['Document', 'H1', 'H2', 'P', 'Link'])
			expect(pdf.raw).toContain(`/Type /StructElem /S /${type} `);
		expect(pdf.raw).toMatch(/\/Type \/OBJR \/Pg \d+ 0 R \/Obj \d+ 0 R/);
		expect(pdf.raw).toMatch(/\/Subtype \/Link [^\n]*\/StructParent \d+/);
		expect(pdf.raw).toContain('/StructParents 0');
		expect(pdf.raw).toContain('/Tabs /S');
		expect(pdf.content).toMatch(
			/\/H1 <<\/MCID 0>> BDC\n(?:[^\n]*\n)?BT [^\n]*\(Intro\) Tj ET[^\n]*\nEMC/,
		);
		expect(pdf.content).toMatch(
			/\/Link <<\/MCID \d+>> BDC\n(?:[^\n]*\n)?BT [^\n]*\(the \) Tj[^\n]*\n(?:[^\n]*\n)?BT [^\n]*\(site ?\) Tj[^\n]*\nEMC\n\/P <<\/MCID/,
		);
		const tabled = inspect((await exportDocxPdf(table.model, { measurer })).bytes);
		for (const type of ['Table', 'TR', 'TD'])
			expect(tabled.raw).toContain(`/Type /StructElem /S /${type} `);
		expect(tabled.content).toContain('/Artifact BMC');
		const plain = inspect((await exportDocxPdf(model, { measurer, taggedPdf: false })).bytes);
		expect(plain.raw).not.toContain('/StructTreeRoot');
		expect(plain.content).not.toContain('BDC');
		// Links work without tags.
		expect(plain.raw).toContain('/Subtype /Link');
	});

	it('marks headers and footers as pagination artifacts', async () => {
		const footer = `<w:ftr xmlns:w="${W}"><w:p>${run('Footer text')}</w:p></w:ftr>`;
		const zip = await docx(
			`<w:p>${run('Body')}</w:p>`,
			{ 'word/footer1.xml': footer },
			`<Relationship Id="rF" Type="${R}/footer" Target="footer1.xml"/>`,
			'<w:footerReference w:type="default" r:id="rF"/>',
		);
		const pdf = inspect((await exportDocxPdf(zip.model, { measurer })).bytes);
		expect(pdf.content).toMatch(
			/\/Artifact <<\/Type \/Pagination \/Subtype \/Footer>> BDC\n[^]*?\(Footer \) Tj[^]*?\(text\) Tj[^\n]*\nEMC$/,
		);
	});

	it('writes PDF/A-2b metadata, an sRGB output intent and the file identifier', async () => {
		const { model } = await docx(`<w:p>${run('Archive')}</w:p>`);
		model.properties = { title: 'Report (final)', creator: 'Ada' };
		const now = new Date(Date.UTC(2026, 0, 2, 3, 4, 5));
		const result = await exportDocxPdf(model, { measurer, pdfA: true, now });
		const pdf = inspect(result.bytes);
		await dump('pdfa', result.bytes);
		expect(pdf.raw.startsWith('%PDF-1.7')).toBe(true);
		expect(pdf.raw).toContain(
			'<pdfaid:part>2</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance>',
		);
		expect(pdf.raw).toContain('<rdf:li xml:lang="x-default">Report (final)</rdf:li>');
		expect(pdf.raw).toContain('<xmp:CreateDate>2026-01-02T03:04:05Z</xmp:CreateDate>');
		expect(pdf.raw).toContain('/CreationDate (D:20260102030405Z)');
		expect(pdf.raw).toMatch(/\/Type \/Metadata \/Subtype \/XML \/Length \d+ >>/);
		expect(pdf.raw).toMatch(/\/OutputIntents \[\d+ 0 R\]/);
		expect(pdf.raw).toContain('/S /GTS_PDFA1 /OutputConditionIdentifier (sRGB IEC61966-2.1)');
		expect(pdf.raw).toMatch(/\/ID \[<[0-9A-F]{32}> <[0-9A-F]{32}>\]/);
		// Standard fonts are not embedded, so the file says it does not conform.
		expect(result.report.pdfA).toEqual({
			conformant: false,
			issues: [expect.stringContaining('Calibri')],
		});
		const embedded = await docx(
			`<w:p><w:r><w:rPr><w:rFonts w:ascii="Synth" w:hAnsi="Synth"/></w:rPr><w:t>AB</w:t></w:r></w:p>`,
		);
		const conformant = await exportDocxPdf(embedded.model, {
			measurer,
			pdfA: true,
			fonts: [{ family: 'Synth', data: syntheticTrueType() }],
		});
		expect(conformant.report.pdfA).toEqual({ conformant: true, issues: [] });
		expect((await exportDocxPdf(model, { measurer })).report.pdfA).toBeUndefined();
	});

	it('shows tracked changes and comments with markup, and hides them without', async () => {
		const { model } = await docx(
			`<w:p>${run('Kept ')}<w:ins w:id="1" w:author="Ada"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="2" w:author="Bob"><w:r><w:delText>gone</w:delText></w:r></w:del><w:commentRangeStart w:id="7"/>${run(' noted')}<w:commentRangeEnd w:id="7"/></w:p>`,
			{
				'word/comments.xml': `<w:comments xmlns:w="${W}"><w:comment w:id="7" w:author="Cy" w:date="2026-03-04T05:06:07Z"><w:p><w:r><w:t>Check this</w:t></w:r></w:p></w:comment></w:comments>`,
			},
			`<Relationship Id="rC" Type="${R}/comments" Target="comments.xml"/>`,
		);
		const marked = await exportDocxPdf(model, { measurer, markup: 'all' });
		await dump('markup', marked.bytes);
		const pdf = inspect(marked.bytes);
		// Ada's insertion and Bob's deletion in their own colours, with a change bar at the margin.
		const colour = (author: string) => {
			let hash = 0;
			for (const c of author) hash = (hash * 31 + c.charCodeAt(0)) >>> 0;
			return hash;
		};
		expect(colour('Ada')).not.toBe(colour('Bob'));
		expect(pdf.content).toMatch(/rg\n[^\n]*\(added\) Tj[^]*? re\nf/);
		expect(pdf.content).toMatch(/\(gone ?\) Tj/);
		expect(pdf.content).toMatch(/\n72 [\d.]+ 1 [\d.]+ re\nf/);
		expect(pdf.raw).toContain('/Subtype /Text');
		expect(pdf.raw).toContain('/Contents (Check this) /T (Cy) /M (D:20260304050607Z)');
		expect(pdf.raw).toContain('/Type /StructElem /S /Annot ');
		const clean = inspect((await exportDocxPdf(model, { measurer })).bytes);
		expect(clean.content).not.toContain('(gone');
		expect(clean.raw).not.toContain('/Subtype /Text');
	});
});

describe('fieldLinkTarget', () => {
	it('reads HYPERLINK, REF, PAGEREF and NOTEREF instructions', () => {
		expect(fieldLinkTarget('HYPERLINK "https://a.test/x"')).toEqual({ uri: 'https://a.test/x' });
		expect(fieldLinkTarget('HYPERLINK \\l "_Toc1"')).toEqual({ anchor: '_Toc1' });
		expect(fieldLinkTarget('HYPERLINK "javascript:alert(1)"')).toBeUndefined();
		expect(fieldLinkTarget(' REF _Ref12 \\h ')).toEqual({ anchor: '_Ref12' });
		expect(fieldLinkTarget('PAGEREF _Toc9 \\h')).toEqual({ anchor: '_Toc9' });
		expect(fieldLinkTarget('REF _Ref12')).toBeUndefined();
	});
});

/** Writes a PDF for manual checks (pypdf) when PDF_DUMP names a folder. */
async function dump(name: string, bytes: Uint8Array) {
	const folder = process.env.PDF_DUMP;
	if (folder) await writeFile(`${folder}/${name}.pdf`, bytes);
}
