import { twips } from './units';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { at, expectParagraph } from './test-support/access';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fixture(bodyXml: string): Promise<{ bytes: Uint8Array }> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`,
	);
	return { bytes: await zip.generateAsync({ type: 'uint8array' }) };
}
async function documentXml(bytes: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	return (await zip.file('word/document.xml')?.async('string')) ?? '';
}

describe('page and column breaks', () => {
	it('preserves pagination cache on no-op saves and invalidates it after text edits', async () => {
		const { bytes } = await fixture(
			'<w:p><w:r><w:lastRenderedPageBreak/><w:t>Cached page</w:t></w:r></w:p>',
		);
		const loaded = await loadDocx(bytes);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.runs).toEqual([{ text: 'Cached page' }]);
		expect(await loaded.save()).toEqual(bytes);
		paragraph.runs[0]!.text += '!';
		const saved = await loaded.save();
		expect(await documentXml(saved)).not.toContain('lastRenderedPageBreak');
		expect((await loadDocx(saved)).model.blocks[0]).toMatchObject({
			runs: [{ text: 'Cached page!' }],
		});
	});
	it('recognizes an authored break beside a calculated marker and retains the authored break', async () => {
		const { bytes } = await fixture(
			'<w:p><w:r><w:lastRenderedPageBreak/><w:br w:type="page"/></w:r></w:p>',
		);
		const loaded = await loadDocx(bytes);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.runs).toEqual([{ text: '', break: 'page' }]);
		paragraph.align = 'center';
		const saved = await loaded.save();
		expect(await documentXml(saved)).toContain('<w:br w:type="page"/>');
		expect(await documentXml(saved)).not.toContain('lastRenderedPageBreak');
	});
	it('keeps guarded rejection for an extended cached marker with unknown attributes', async () => {
		const { bytes } = await fixture(
			'<w:p><w:r><w:lastRenderedPageBreak w:unknown="keep"/><w:t>Text</w:t></w:r></w:p>',
		);
		const loaded = await loadDocx(bytes);
		expectParagraph(loaded.model.blocks[0]).runs[0]!.text += '!';
		await expect(loaded.save()).rejects.toThrow('cannot safely relocate');
		expect(await documentXml(bytes)).toContain('w:unknown="keep"');
	});
	it('models a page break run distinctly from a line break and edits it safely', async () => {
		const { bytes } = await fixture(
			'<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>After</w:t></w:r></w:p>',
		);
		const loaded = await loadDocx(bytes);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.runs).toMatchObject([
			{ text: 'Before' },
			{ text: '', break: 'page' },
			{ text: 'After' },
		]);
		expect(loaded.model.warnings).not.toContain(
			expect.stringContaining('non-line breaks other than page and column breaks'),
		);
		paragraph.align = 'center';
		const saved = await loaded.save();
		expect(saved).not.toEqual(bytes);
		const xml = await documentXml(saved);
		expect(xml).toContain('<w:br w:type="page"/>');
		const reopened = await loadDocx(saved);
		expect(reopened.model.blocks[0]).toMatchObject({
			align: 'center',
			runs: [{ text: 'Before' }, { text: '', break: 'page' }, { text: 'After' }],
		});
	});

	it('models a column break run and round-trips it', async () => {
		const { bytes } = await fixture('<w:p><w:r><w:br w:type="column"/></w:r></w:p>');
		const loaded = await loadDocx(bytes);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.runs).toMatchObject([{ text: '', break: 'column' }]);
		paragraph.spacingBeforeTwips = twips(240);
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain('<w:br w:type="column"/>');
	});

	it('parses and writes pageBreakBefore', async () => {
		const { bytes } = await fixture(
			'<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>',
		);
		const loaded = await loadDocx(bytes);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.pageBreakBefore).toBe(true);
		paragraph.pageBreakBefore = false;
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain('<w:pageBreakBefore w:val="0"/>');
		delete paragraph.pageBreakBefore;
		expect(await documentXml(await loaded.save())).not.toContain('pageBreakBefore');
	});
});

describe('footnote and endnote reference marks', () => {
	it('models a footnote reference run and keeps it when its paragraph is edited', async () => {
		const { bytes } = await fixture(
			'<w:p><w:r><w:t>See</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="3"/></w:r></w:p>',
		);
		const loaded = await loadDocx(bytes);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.runs).toMatchObject([
			{ text: 'See' },
			{ text: '', noteReference: { kind: 'footnote', id: '3' } },
		]);
		expect(await loaded.save()).toEqual(bytes);
		paragraph.align = 'center';
		at(paragraph.runs, 0).text = 'See also';
		const saved = await loaded.save();
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).toContain('See also');
		expect(xml).toMatch(
			/<w:r><w:rPr><w:vertAlign w:val="superscript"\/><\/w:rPr><w:footnoteReference w:id="3"\/><\/w:r>/,
		);
	});

	it('models an endnote reference run', async () => {
		const { bytes } = await fixture('<w:p><w:r><w:endnoteReference w:id="1"/></w:r></w:p>');
		const loaded = await loadDocx(bytes);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.runs).toMatchObject([
			{ text: '', noteReference: { kind: 'endnote', id: '1' } },
		]);
	});
});
