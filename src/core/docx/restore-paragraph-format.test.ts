import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { rejectRevision } from './revision-commands';
import { restoreParagraphFormatting } from './restore-paragraph-format';
import { expectParagraph } from './test-support/access';
import type { Paragraph } from './model';
import { twips } from './units';

describe('complete paragraph formatting restoration', () => {
	it('retains opaque properties, text, bookmarks and paragraph-mark history while restoring formatting', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:paragraph-test"><w:body><w:p><w:pPr><w:jc w:val="center"/><w:keepNext/><w:rPr><w:ins w:id="mark" w:author="Grace"/></w:rPr><w:pPrChange w:id="7" w:author="Ada"><w:pPr><w:pStyle w:val="Original"/><w:bidi/><w:jc w:val="end"/><w:tabs><w:tab w:val="right" w:pos="1440"/></w:tabs><x:extension x:value="preserved"/></w:pPr></w:pPrChange></w:pPr><w:bookmarkStart w:id="3" w:name="Anchor"/><w:r><w:t>Text</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const before = structuredClone(loaded.model);
		const model = rejectRevision(loaded.model, '7');
		expect(loaded.model).toEqual(before);
		const paragraph = expectParagraph(model.blocks[0]);
		expect(paragraph).toMatchObject({
			style: 'Original',
			direction: 'rtl',
			justification: 'end',
			tabStops: [{ align: 'right', posTwips: 1440 }],
			bookmarks: ['Anchor'],
			markRevision: { id: 'mark', author: 'Grace' },
			runs: [{ text: 'Text' }],
		});
		expect(paragraph.keepNext).toBeUndefined();
		paragraph.spacingAfterTwips = twips(120);
		paragraph.runs[0]!.text += '!';
		const bytes = await loaded.save(model);
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('x:value="preserved"');
		expect(xml).not.toContain('pPrChange');
		const actual = expectParagraph((await loadDocx(bytes)).model.blocks[0]);
		expect(actual).toMatchObject({
			spacingAfterTwips: 120,
			justification: 'end',
			direction: 'rtl',
			bookmarks: ['Anchor'],
			markRevision: { author: 'Grace' },
			runs: [{ text: 'Text!' }],
		});
	});
	it('leaves a paragraph intact when its snapshot is absent or invalid', () => {
		for (const previousParagraphPropertiesXml of [
			undefined,
			'<pPr xmlns="urn:wrong"/>',
			'<!DOCTYPE pPr><pPr/>',
		]) {
			const paragraph: Paragraph = {
				type: 'paragraph',
				id: 'p',
				runs: [{ text: 'Text' }],
				align: 'center',
				formatRevision: {
					kind: 'paragraphChange',
					id: '7',
					author: 'Ada',
					...(previousParagraphPropertiesXml ? { previousParagraphPropertiesXml } : {}),
				},
			};
			const before = structuredClone(paragraph);
			expect(() => restoreParagraphFormatting(paragraph)).toThrow();
			expect(paragraph).toEqual(before);
		}
	});
});
