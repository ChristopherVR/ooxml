import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 20 };
const noteText = (result: ReturnType<typeof layoutDocumentModel>) =>
	result.pages.flatMap((page) =>
		(page.footnotes ?? []).flatMap((note) =>
			note.paragraphs.flatMap((paragraph) =>
				paragraph.lines.map((line) => line.fragments.map((fragment) => fragment.text).join('')),
			),
		),
	);

describe('continued footnotes through DOCX saving', () => {
	it('keeps a late reference attached to its displayed line after a body edit and save/reopen', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${w}"><w:body><w:p><w:pPr><w:widowControl w:val="0"/></w:pPr><w:r><w:t>Claim</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="3000" w:h="1500"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"/></w:sectPr></w:body></w:document>`,
		);
		zip.file(
			'word/footnotes.xml',
			`<w:footnotes xmlns:w="${w}"><w:footnote w:id="2"><w:p><w:r><w:t>Source</w:t></w:r></w:p></w:footnote></w:footnotes>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected body paragraph');
		paragraph.runs[0]!.text = 'A\nB\nC\nD\nE\nClaim';
		const before = layoutDocumentModel(loaded.model, measurer);
		expect(before.pages[0]!.footnotes).toBeUndefined();
		expect(before.pages[1]!.footnotes![0]!.id).toBe('2');
		const reopened = await loadDocx(await loaded.save(loaded.model));
		const after = layoutDocumentModel(reopened.model, measurer);
		expect(after.pages.map((page) => page.footnotes?.map((note) => note.id))).toEqual(
			before.pages.map((page) => page.footnotes?.map((note) => note.id)),
		);
	});
	it('paginates after editing and preserves the original note, numbering mark and custom separator XML on reopen', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>Claim</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="3000" w:h="1500"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"/></w:sectPr></w:body></w:document>`,
		);
		zip.file(
			'word/footnotes.xml',
			`<w:footnotes xmlns:w="${w}"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:t>Continued:</w:t><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Source</w:t></w:r></w:p></w:footnote></w:footnotes>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const edited = structuredClone(loaded.model);
		const note = edited.footnotes![0]!;
		const paragraph = note.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected note paragraph');
		paragraph.runs[1]!.text = Array.from({ length: 12 }, (_, index) => `Line${index}`).join('\n');
		const before = structuredClone(edited);
		const layout = layoutDocumentModel(edited, measurer);
		expect(layout.pages.length).toBeGreaterThan(2);
		expect(edited).toEqual(before);
		const saved = await loaded.save(edited);
		const xml = await (await JSZip.loadAsync(saved)).file('word/footnotes.xml')!.async('string');
		expect(xml).toContain('Continued:');
		expect(xml).toContain('<w:continuationSeparator/>');
		expect(xml).toContain('<w:footnoteRef/>');
		const reopened = await loadDocx(saved);
		expect(reopened.model.footnotes).toHaveLength(1);
		expect(reopened.model.footnotes![0]!.id).toBe('2');
		const after = layoutDocumentModel(reopened.model, measurer);
		expect(noteText(after)).toEqual(noteText(layout));
		expect(after.pages.map((page) => page.footnotes?.map((note) => note.continued))).toEqual(
			layout.pages.map((page) => page.footnotes?.map((note) => note.continued)),
		);
	});
});
