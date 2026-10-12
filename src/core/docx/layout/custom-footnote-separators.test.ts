import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, saveDocx } from '../index';
import { layoutDocumentModel } from './layout';
import { layoutSections } from './page-flow';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 20 };
const paragraph = (text: string) =>
	`<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
async function fixture(separator: string, continuation = paragraph('Continued sources')) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>Claim</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="4500" w:h="1800"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/footnotes.xml',
		`<w:footnotes xmlns:w="${w}"><w:footnote w:type="separator" w:id="-1">${separator}</w:footnote><w:footnote w:type="continuationSeparator" w:id="0">${continuation}</w:footnote><w:footnote w:id="2">${Array.from({ length: 12 }, (_, index) => paragraph('Source ' + index)).join('')}</w:footnote></w:footnotes>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const textOf = (page: ReturnType<typeof layoutDocumentModel>['pages'][number]) =>
	page.footnoteSeparator?.paragraphs
		.flatMap((p) => p.lines.flatMap((line) => line.fragments.map((fragment) => fragment.text)))
		.join('');

describe('custom footnote separators', () => {
	it('reports page-wide multi-column fallback honestly', () => {
		const result = layoutSections(
			{
				sections: [
					{
						page: {
							widthPx: 300,
							heightPx: 150,
							marginTopPx: 0,
							marginBottomPx: 0,
							marginLeftPx: 0,
							marginRightPx: 0,
						},
						columns: { count: 2, gapPx: 10 },
						blocks: [
							{
								kind: 'paragraph',
								id: 'body',
								runs: [{ text: 'Claim' }],
								footnotes: [
									{
										id: '2',
										paragraphs: [{ kind: 'paragraph', id: 'note', runs: [{ text: 'Source' }] }],
									},
								],
							},
						],
					},
				],
			},
			measurer,
		);
		expect(result.approximations).toContain(
			'Multi-column footnotes use a page-wide area; native per-column footnote balancing is not yet modeled.',
		);
	});
	it('parses mixed text/rule runs separately without numbering separator stories', async () => {
		const loaded = await fixture(
			'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Sources</w:t><w:separator/><w:t> below</w:t></w:r></w:p>',
		);
		expect(loaded.model.footnotes).toHaveLength(1);
		const separator = loaded.model.footnoteSeparators!.separator!.blocks[0]!;
		expect(separator).toMatchObject({
			runs: [
				{ text: 'Sources', bold: true },
				{ text: '', noteSeparator: 'separator', bold: true },
				{ text: ' below', bold: true },
			],
		});
		const result = layoutDocumentModel(loaded.model, measurer);
		expect(textOf(result.pages[0]!)).toBe('Sources below');
		expect(
			result.pages[0]!.footnoteSeparator!.paragraphs.flatMap((p) =>
				p.lines.flatMap((line) => line.fragments),
			).find((fragment) => fragment.object)?.object,
		).toMatchObject({ noteSeparator: 'separator', widthPx: 192 });
	});
	it('reserves multiline custom stories and selects continuation content on later pages', async () => {
		const loaded = await fixture(paragraph('Sources') + paragraph('References'));
		const result = layoutDocumentModel(loaded.model, measurer);
		expect(result.pages[0]!.footnoteSeparator!.heightPx).toBe(40);
		expect(textOf(result.pages[0]!)).toBe('SourcesReferences');
		expect(textOf(result.pages[1]!)).toBe('Continued sources');
		for (const page of result.pages) {
			const height =
				page.footnoteSeparator!.heightPx +
				page.footnotes!.reduce((sum, note) => sum + note.heightPx, 0);
			const bottom = Math.max(
				0,
				...page.columns.flatMap((column) =>
					column.blocks.map((block) => block.yPx + block.heightPx),
				),
			);
			expect(bottom + height).toBeLessThanOrEqual(120);
		}
	});
	it('supports explicitly empty separator stories and full-width continuation marks', async () => {
		const loaded = await fixture('', '<w:p><w:r><w:continuationSeparator/></w:r></w:p>');
		const result = layoutDocumentModel(loaded.model, measurer);
		expect(result.pages[0]!.footnoteSeparator).toEqual({ heightPx: 0, paragraphs: [] });
		const rule = result.pages[1]!.footnoteSeparator!.paragraphs[0]!.lines[0]!.fragments[0]!.object!;
		expect(rule).toMatchObject({ noteSeparator: 'continuationSeparator', widthPx: 300 });
		expect(result.pages[1]!.footnoteSeparator!.heightPx).toBe(20);
	});
	it('preserves imported separator XML on unrelated editing and saves edited separator stories', async () => {
		const loaded = await fixture(paragraph('Sources'));
		const before = (await JSZip.loadAsync(await loaded.save(loaded.model))).file(
			'word/footnotes.xml',
		)!;
		const original = await before.async('string');
		const body = loaded.model.blocks[0]!;
		if (body.type === 'paragraph') body.runs[0]!.text = 'Edited claim';
		const unchanged = await JSZip.loadAsync(await loaded.save(loaded.model));
		expect(await unchanged.file('word/footnotes.xml')!.async('string')).toBe(original);
		const separator = loaded.model.footnoteSeparators!.separator!.blocks[0]!;
		if (separator.type !== 'paragraph') throw new Error('Expected paragraph');
		separator.runs[0]!.text = 'Evidence';
		const reopened = await loadDocx(await loaded.save(loaded.model));
		expect(textOf(layoutDocumentModel(reopened.model, measurer).pages[0]!)).toBe('Evidence');
		expect(reopened.model.footnotes).toEqual(loaded.model.footnotes);
	});
	it('writes custom stories and rule marks from a standalone model', async () => {
		const loaded = await fixture('<w:p><w:r><w:separator/></w:r></w:p>');
		const reopened = await loadDocx(await saveDocx(loaded.model));
		expect(reopened.model.footnoteSeparators!.separator!.blocks[0]).toMatchObject({
			runs: [{ noteSeparator: 'separator' }],
		});
		expect(reopened.model.footnotes).toHaveLength(1);
	});
});
