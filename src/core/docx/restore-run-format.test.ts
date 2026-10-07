import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { rejectRevision } from './revision-commands';
import { expectParagraph } from './test-support/access';
import { restoreRunFormatting } from './restore-run-format';
import type { TextRun } from './model';

describe('restoring complete run properties', () => {
	it('restores language, style and opaque properties without retaining newly applied formatting', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:format-test"><w:body><w:p><w:r><w:rPr><w:b/><w:i/><w:lang w:val="fr-FR"/><w:rtl/><w:rPrChange w:id="7" w:author="Ada"><w:rPr><w:rStyle w:val="Original"/><w:b w:val="0"/><w:szCs w:val="28"/><w:lang w:val="en-GB" w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:rtl w:val="0"/><x:extension x:value="preserved"/></w:rPr></w:rPrChange></w:rPr><w:t>Text</w:t></w:r></w:p></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const original = structuredClone(loaded.model);
		const rejected = rejectRevision(loaded.model, '7');
		expect(loaded.model).toEqual(original);
		const run = expectParagraph(rejected.blocks[0]).runs[0]!;
		expect(run).toMatchObject({
			text: 'Text',
			style: 'Original',
			bold: false,
			language: 'en-GB',
			eastAsiaLanguage: 'ja-JP',
			bidiLanguage: 'ar-SA',
			rtl: false,
		});
		expect(run.italic).toBeUndefined();
		run.text = 'Edited';
		run.bold = true;
		const bytes = await loaded.save(rejected);
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:szCs w:val="28"');
		expect(xml).toContain('x:value="preserved"');
		expect(xml).not.toContain('rPrChange');
		const actual = expectParagraph((await loadDocx(bytes)).model.blocks[0]).runs[0]!;
		expect(actual).toMatchObject({
			text: 'Edited',
			bold: true,
			style: 'Original',
			language: 'en-GB',
			rtl: false,
		});
		expect(actual.italic).toBeUndefined();
	});
	it('leaves the run intact when a prior snapshot is absent or invalid', () => {
		for (const previousRunPropertiesXml of [
			undefined,
			'<rPr xmlns="urn:wrong"/>',
			'<!DOCTYPE rPr><rPr/>',
		]) {
			const run: TextRun = {
				text: 'Text',
				bold: true,
				revision: {
					kind: 'formatChange',
					id: '7',
					author: 'Ada',
					...(previousRunPropertiesXml ? { previousRunPropertiesXml } : {}),
				},
			};
			const original = structuredClone(run);
			expect(() => restoreRunFormatting(run)).toThrow();
			expect(run).toEqual(original);
		}
	});
});
