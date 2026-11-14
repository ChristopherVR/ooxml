import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { convertNotes, loadDocx, type DocumentModel, type Paragraph } from './index';
import { at, expectParagraph } from './test-support/access';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const PNG = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
		'base64',
	),
);
const ns = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}"`;
const reference = (kind: string, id: string, style: string) =>
	`<w:r><w:rPr><w:rStyle w:val="${style}"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r>`;
const note = (kind: 'footnote' | 'endnote', id: string, body: string) => {
	const cap = kind === 'footnote' ? 'Footnote' : 'Endnote';
	return `<w:${kind} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="${cap}Text"/></w:pPr><w:r><w:rPr><w:rStyle w:val="${cap}Reference"/></w:rPr><w:${kind}Ref/></w:r>${body}</w:p></w:${kind}>`;
};
const separators = (kind: string) =>
	`<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;
const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const picture = `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="7" name="Picture 7"/><a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const style = (type: string, id: string) =>
	`<w:style w:type="${type}" w:styleId="${id}"><w:name w:val="${id}"/></w:style>`;

/** Shaped like Word 16's own output: two footnotes, one endnote, both note style sets. */
async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${ns}><w:body><w:p>${text('Alpha claim.')}${reference('footnote', '1', 'FootnoteReference')}${text(' Beta claim.')}${reference('endnote', '1', 'EndnoteReference')}</w:p><w:tbl><w:tr><w:tc><w:p>${text('Gamma claim.')}${reference('footnote', '2', 'FootnoteReference')}</w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/footnotes.xml',
		`<w:footnotes ${ns}>${separators('footnote')}${note('footnote', '1', `${text(' First foot, ')}<w:hyperlink r:id="rId8">${text('source')}</w:hyperlink>${picture}`)}${note('footnote', '2', text(' Second foot.'))}</w:footnotes>`,
	);
	zip.file(
		'word/_rels/footnotes.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId8" Type="${R}/hyperlink" Target="https://example.com/source" TargetMode="External"/><Relationship Id="rId9" Type="${R}/image" Target="media/image1.png"/></Relationships>`,
	);
	zip.file('word/media/image1.png', PNG);
	zip.file(
		'word/endnotes.xml',
		`<w:endnotes ${ns}>${separators('endnote')}${note('endnote', '1', text(' Only end.'))}</w:endnotes>`,
	);
	zip.file(
		'word/styles.xml',
		`<w:styles ${ns}>${style('paragraph', 'FootnoteText')}${style('character', 'FootnoteReference')}${style('paragraph', 'EndnoteText')}${style('character', 'EndnoteReference')}</w:styles>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

const references = (model: DocumentModel) => {
	const found: string[] = [];
	const visit = (paragraph: Paragraph) => {
		for (const run of paragraph.runs)
			if (run.noteReference)
				found.push(`${run.noteReference.kind}:${run.noteReference.id}:${run.style ?? ''}`);
	};
	for (const block of model.blocks)
		if (block.type === 'paragraph') visit(block);
		else for (const row of block.rows) for (const cell of row) cell.paragraphs.forEach(visit);
	return found;
};
const noteText = (blocks: { blocks: unknown[] }) =>
	expectParagraph(blocks.blocks[0] as Paragraph)
		.runs.map((run) => run.text)
		.join('');

describe('convertNotes', () => {
	it('converts every footnote to an endnote, keeping links and pictures', async () => {
		const loaded = await loadDocx(await fixture());
		const { model, plan } = convertNotes(loaded.model, 'footnotesToEndnotes');
		expect(plan.skipped).toEqual([]);
		expect(model.footnotes).toEqual([]);
		expect(model.endnotes!.map((item) => item.id)).toEqual(['1', '2', '3']);
		expect(references(model)).toEqual([
			'endnote:2:EndnoteReference',
			'endnote:1:EndnoteReference',
			'endnote:3:EndnoteReference',
		]);
		const moved = at(model.endnotes, 1);
		const paragraph = expectParagraph(moved.blocks[0]);
		expect(paragraph.style).toBe('EndnoteText');
		expect(paragraph.id.startsWith('en2-')).toBe(true);
		expect(paragraph.runs[0]).toMatchObject({ noteMark: 'endnote', style: 'EndnoteReference' });

		const bytes = await loaded.save(model);
		const zip = await JSZip.loadAsync(bytes);
		const footnotes = await zip.file('word/footnotes.xml')!.async('string');
		expect(footnotes).not.toMatch(/<w:footnote w:id="[12]"/);
		expect(footnotes).toContain('<w:separator/>');
		const endnotes = await zip.file('word/endnotes.xml')!.async('string');
		expect(endnotes).toMatch(
			/<w:endnote w:id="2"><w:p><w:pPr><w:pStyle w:val="EndnoteText"\/><\/w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"\/><\/w:rPr><w:endnoteRef\/>/,
		);
		const rels = await zip.file('word/_rels/endnotes.xml.rels')!.async('string');
		expect(rels).toContain('Target="https://example.com/source"');
		expect(rels).toContain('Target="media/image1.png"');
		const document = await zip.file('word/document.xml')!.async('string');
		expect(document).not.toContain('footnoteReference');
		expect(document).toContain(
			'<w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="2"/>',
		);

		const reopened = (await loadDocx(bytes)).model;
		expect(reopened.footnotes).toEqual([]);
		expect(reopened.endnotes!.map(noteText)).toEqual([
			' Only end.',
			' First foot, source',
			' Second foot.',
		]);
		const runs = expectParagraph(at(reopened.endnotes, 1).blocks[0]).runs;
		expect(runs.find((run) => run.link)?.link?.href).toBe('https://example.com/source');
		expect(runs.find((run) => run.image)?.image?.partName).toBe('word/media/image1.png');
		expect(references(reopened)).toEqual(references(model));
	});

	it('swaps footnotes and endnotes', async () => {
		const loaded = await loadDocx(await fixture());
		const { model } = convertNotes(loaded.model, 'swap');
		expect(model.footnotes!.map(noteText)).toEqual([' Only end.']);
		expect(model.endnotes!.map(noteText)).toEqual([' First foot, source', ' Second foot.']);
		expect(references(model)).toEqual([
			'endnote:2:EndnoteReference',
			'footnote:3:FootnoteReference',
			'endnote:3:EndnoteReference',
		]);
		const reopened = (await loadDocx(await loaded.save(model))).model;
		expect(reopened.footnotes!.map(noteText)).toEqual([' Only end.']);
		expect(reopened.endnotes!.map(noteText)).toEqual([' First foot, source', ' Second foot.']);
		expect(references(reopened)).toEqual(references(model));
	});

	it('converts one note and leaves the others', async () => {
		const loaded = await loadDocx(await fixture());
		const { model } = convertNotes(loaded.model, { kind: 'footnote', id: '1' });
		expect(model.footnotes!.map((item) => item.id)).toEqual(['2']);
		expect(model.endnotes!.map(noteText)).toEqual([' Only end.', ' First foot, source']);
		expect(references(model)).toEqual([
			'endnote:2:EndnoteReference',
			'endnote:1:EndnoteReference',
			'footnote:2:FootnoteReference',
		]);
		const back = convertNotes(model, { kind: 'endnote', id: '2' }, { endnote: ['1'] }).model;
		expect(references(back)[0]).toBe('footnote:3:FootnoteReference');
		const reopened = (await loadDocx(await loaded.save(back))).model;
		expect(reopened.footnotes!.map(noteText)).toEqual([' Second foot.', ' First foot, source']);
		expect(reopened.endnotes!.map(noteText)).toEqual([' Only end.']);
	});

	it('keeps the source styles when the document lacks the counterparts, and never reuses reserved ids', async () => {
		const loaded = await loadDocx(await fixture());
		const bare: DocumentModel = { ...loaded.model };
		delete bare.paragraphStyles;
		delete bare.characterStyles;
		const { model } = convertNotes(bare, 'endnotesToFootnotes', { footnote: ['7'] });
		const moved = at(model.footnotes, 2);
		expect(moved.id).toBe('8');
		expect(expectParagraph(moved.blocks[0]).style).toBe('EndnoteText');
		expect(references(model)[1]).toBe('footnote:8:EndnoteReference');
		expect(model.endnotes).toEqual([]);
	});

	it('leaves notes with source-only drawings in place and reports them', async () => {
		const loaded = await loadDocx(await fixture());
		const image = expectParagraph(at(loaded.model.footnotes, 0).blocks[0]).runs.find(
			(run) => run.image,
		)!.image!;
		image.unsupported = 'Chart';
		const { model, plan } = convertNotes(loaded.model, 'footnotesToEndnotes');
		expect(plan.skipped).toEqual([{ kind: 'footnote', id: '1' }]);
		expect(model.footnotes!.map((item) => item.id)).toEqual(['1']);
		expect(references(model)[0]).toBe('footnote:1:FootnoteReference');
	});
});
