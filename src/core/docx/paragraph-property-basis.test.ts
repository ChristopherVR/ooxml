import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, saveDocx, rejectAllRevisions } from './index';
import { expectParagraph } from './test-support/access';
import { first, buildXml, parseXml, WORD_NS } from './xml';
import { paragraphAttrs, paragraphFromAttrs } from './ui/paragraph-attributes';
import { twips } from './units';

describe('complete paragraph property bases', () => {
	it.each(['alignment', 'spacing', 'indent', 'multiple'])(
		'retains native paragraph-mark fonts when rejecting %s formatting',
		async (name) => {
			const fixture = (state: string) =>
				new URL(
					`./__fixtures__/review-paragraph-formatting/${name}-${state}.docx`,
					import.meta.url,
				);
			const loaded = await loadDocx(new Uint8Array(await readFile(fixture('tracked'))));
			const native = await loadDocx(new Uint8Array(await readFile(fixture('rejected'))));
			const expected = first(
				parseXml(expectParagraph(native.model.blocks[0]).sourceParagraphPropertiesXml!)
					.documentElement,
				'rPr',
			)!;
			const model = rejectAllRevisions(loaded.model);
			for (const bytes of [await loaded.save(model), await saveDocx(model)]) {
				const reopened = await loadDocx(bytes);
				const actual = first(
					parseXml(expectParagraph(reopened.model.blocks[0]).sourceParagraphPropertiesXml!)
						.documentElement,
					'rPr',
				)!;
				expect(buildXml(actual)).toBe(buildXml(expected));
			}
		},
	);
	it('preserves opaque properties, paragraph-mark formatting and section metadata after editor conversion and known edits', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${WORD_NS}" xmlns:x="urn:paragraph-basis"><w:body><w:p><w:pPr x:flag="kept"><w:pBdr><w:top w:val="single" w:sz="8" x:detail="border"/></w:pBdr><w:spacing w:after="100" w:beforeAutospacing="1"/><w:rPr><w:rFonts w:ascii="Arial" w:hint="eastAsia"/></w:rPr><w:sectPr x:detail="section"><w:pgSz w:w="8000" w:h="9000"/><w:cols w:num="2"/></w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12000" w:h="15000"/></w:sectPr></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const model = structuredClone(loaded.model);
		const original = expectParagraph(model.blocks[0]);
		model.blocks[0] = paragraphFromAttrs(paragraphAttrs(original), original.id, [
			{ text: 'Edited' },
		]);
		expectParagraph(model.blocks[0]).spacingAfterTwips = twips(200);
		for (const bytes of [await loaded.save(model), await saveDocx(model)]) {
			const reopened = await loadDocx(bytes);
			const paragraph = expectParagraph(reopened.model.blocks[0]);
			const properties = parseXml(paragraph.sourceParagraphPropertiesXml!).documentElement;
			expect(properties.getAttributeNS('urn:paragraph-basis', 'flag')).toBe('kept');
			expect(first(properties, 'spacing')!.getAttributeNS(WORD_NS, 'beforeAutospacing')).toBe('1');
			expect(
				first(first(properties, 'pBdr'), 'top')!.getAttributeNS('urn:paragraph-basis', 'detail'),
			).toBe('border');
			expect(first(first(properties, 'rPr'), 'rFonts')!.getAttributeNS(WORD_NS, 'hint')).toBe(
				'eastAsia',
			);
			expect(first(properties, 'sectPr')!.getAttributeNS('urn:paragraph-basis', 'detail')).toBe(
				'section',
			);
			expect(reopened.model.sections![0]!.pageWidthTwips).toBe(8000);
			expect(paragraph.spacingAfterTwips).toBe(200);
		}
	});
});
