import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx } from './index';
import { sectionsOf, sectionLayoutJson, sectionsFromLayout } from './section-layout';
const fixture = () =>
	new Uint8Array(
		readFileSync(
			new URL('./__fixtures__/footnote-columns/two-two-custom-eighty.docx', import.meta.url),
		),
	);
describe('footnote column editing and preservation', () => {
	it('preserves native parts byte-for-byte on an unchanged save', async () => {
		const bytes = fixture(),
			loaded = await loadDocx(bytes);
		const original = await JSZip.loadAsync(bytes),
			saved = await JSZip.loadAsync(await loaded.save(loaded.model));
		for (const path of ['word/document.xml', 'word/footnotes.xml', 'word/settings.xml'])
			expect(await saved.file(path)!.async('string')).toBe(
				await original.file(path)!.async('string'),
			);
	});
	it('changes columns without losing separator stories or other compatibility settings', async () => {
		const loaded = await loadDocx(fixture());
		const model = structuredClone(loaded.model);
		model.sections![0]!.footnoteColumns = 3;
		model.compatibilityMode = 14;
		const saved = await loaded.save(model),
			reopened = await loadDocx(saved);
		expect(reopened.model.sections![0]!.footnoteColumns).toBe(3);
		expect(reopened.model.compatibilityMode).toBe(14);
		expect(reopened.model.footnoteSeparators).toEqual(loaded.model.footnoteSeparators);
		const original = await JSZip.loadAsync(fixture()),
			zip = await JSZip.loadAsync(saved);
		expect(await zip.file('word/footnotes.xml')!.async('string')).toBe(
			await original.file('word/footnotes.xml')!.async('string'),
		);
		expect(await zip.file('word/settings.xml')!.async('string')).toContain('compatibilityMode');
	});
	it('saves independent columns and compatibility mode on a new document and history layout', async () => {
		const model = createDocument();
		model.compatibilityMode = 15;
		model.sections = sectionsOf(model);
		model.sections[0]!.footnoteColumns = 46;
		const history = sectionLayoutJson(model.sections);
		expect(sectionsFromLayout(history, model.sections, model.blocks)[0]!.footnoteColumns).toBe(46);
		const reopened = await loadDocx(await saveDocx(model));
		expect(reopened.model.compatibilityMode).toBe(15);
		expect(reopened.model.sections![0]!.footnoteColumns).toBe(46);
	});
	it('declares extension prefixes without overwriting imported namespace bindings', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="urn:custom" xmlns:mc="urn:custom-mc"><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr><w15:payload value="keep"/></w:sectPr></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const next = structuredClone(loaded.model);
		next.sections![0]!.footnoteColumns = 2;
		const bytes = await loaded.save(next),
			saved = await JSZip.loadAsync(bytes);
		const xml = await saved.file('word/document.xml')!.async('string');
		expect(xml).toContain('xmlns:w15="urn:custom"');
		expect(xml).toContain('xmlns:mc="urn:custom-mc"');
		expect(xml).toContain('<w15:payload value="keep"/>');
		expect(xml).toContain('mc1:Ignorable="w151"');
		expect((await loadDocx(bytes)).model.sections![0]!.footnoteColumns).toBe(2);
	});
	it.each([12, 14])(
		'preserves mode %s and both native separator parts while editing body text',
		async (mode) => {
			const bytes = new Uint8Array(
				readFileSync(
					new URL(
						'./__fixtures__/footnote-columns/mode' + mode + '-two-custom-eighty.docx',
						import.meta.url,
					),
				),
			);
			const loaded = await loadDocx(bytes),
				next = structuredClone(loaded.model);
			const paragraph = next.blocks.find((block) => block.type === 'paragraph')!;
			if (paragraph.type !== 'paragraph') throw new Error('Missing paragraph');
			paragraph.runs[0]!.text = 'Prefix ' + paragraph.runs[0]!.text;
			const savedBytes = await loaded.save(next),
				saved = await JSZip.loadAsync(savedBytes),
				original = await JSZip.loadAsync(bytes);
			for (const path of ['word/footnotes.xml', 'word/settings.xml'])
				expect(await saved.file(path)!.async('string')).toBe(
					await original.file(path)!.async('string'),
				);
			const reopened = await loadDocx(savedBytes);
			expect(reopened.model.compatibilityMode).toBe(mode);
			expect(reopened.model.footnoteSeparators).toEqual(loaded.model.footnoteSeparators);
			expect(reopened.model.footnotes).toEqual(loaded.model.footnotes);
		},
	);
});
