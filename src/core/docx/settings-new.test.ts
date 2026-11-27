import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { saveDocx } from './save';
import type { DocumentModel } from './model';
import { twips } from './units';

const page = {
	width: 816,
	height: 1056,
	marginTop: 96,
	marginRight: 96,
	marginBottom: 96,
	marginLeft: 96,
};
const model = (extra: Partial<DocumentModel> = {}): DocumentModel => ({
	blocks: [{ type: 'paragraph', id: 'p', runs: [{ text: 'x' }] }],
	page,
	warnings: [],
	...extra,
});

describe('settings part of a new document', () => {
	it('writes a whole settings part with its relationship and content type', async () => {
		const output = await JSZip.loadAsync(await saveDocx(model()));
		const settings = await output.file('word/settings.xml')?.async('string');
		expect(settings).toMatch(
			/<w:settings [^>]*><w:zoom w:percent="100"\/><w:characterSpacingControl w:val="doNotCompress"\/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http:\/\/schemas.microsoft.com\/office\/word" w:val="15"\/><\/w:compat><\/w:settings>/,
		);
		expect(await output.file('[Content_Types].xml')?.async('string')).toContain(
			'/word/settings.xml',
		);
		expect(await output.file('word/_rels/document.xml.rels')?.async('string')).toContain(
			'relationships/settings',
		);
	});

	it('keeps the model mode and tab stop, and the single-setting writers patch the part in order', async () => {
		const bytes = await saveDocx(
			model({ compatibilityMode: 14, defaultTabStopTwips: twips(360), evenAndOddHeaders: true }),
		);
		const settings = await (
			await JSZip.loadAsync(bytes)
		)
			.file('word/settings.xml')
			?.async('string');
		expect(settings).toContain('<w:defaultTabStop w:val="360"/>');
		expect(settings).toContain('w:val="14"');
		expect(settings).not.toContain('w:val="15"');
		expect(settings?.indexOf('evenAndOddHeaders')).toBeGreaterThan(
			settings?.indexOf('defaultTabStop') ?? Infinity,
		);
		expect(settings?.indexOf('evenAndOddHeaders')).toBeLessThan(
			settings?.indexOf('w:compat>') ?? -1,
		);
		const reopened = (await loadDocx(bytes)).model;
		expect(reopened.compatibilityMode).toBe(14);
		expect(reopened.defaultTabStopTwips).toBe(360);
		expect(reopened.evenAndOddHeaders).toBe(true);
	});

	it('declares the separator notes when the document has notes', async () => {
		const bytes = await saveDocx(
			model({
				blocks: [
					{
						type: 'paragraph',
						id: 'p',
						runs: [{ text: 'x' }, { text: '', noteReference: { kind: 'footnote', id: '1' } }],
					},
				],
				footnotes: [
					{
						id: '1',
						blocks: [
							{
								type: 'paragraph',
								id: 'fn1-p0',
								runs: [{ text: '', noteMark: 'footnote' }, { text: ' Note' }],
							},
						],
					},
				],
			}),
		);
		const settings = await (
			await JSZip.loadAsync(bytes)
		)
			.file('word/settings.xml')
			?.async('string');
		expect(settings).toContain(
			'<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>',
		);
		expect(settings).not.toContain('w:endnotePr');
	});
});
