import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { glossaryBlocksWritable, glossaryEntryKey, glossaryPreview, loadDocx } from './index';
import { createDocument } from './model';
import { saveDocx } from './save';
import { must } from './test-support/access';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('glossary building blocks', () => {
	it('creates the glossary part for a saved selection and reads it back', async () => {
		const model = createDocument();
		model.glossary = [
			{
				name: 'Sign-off',
				gallery: 'docParts',
				category: 'General',
				description: 'Closing line',
				blocks: [
					{ type: 'paragraph', id: 'a', runs: [{ text: 'Kind regards,', bold: true }] },
					{ type: 'paragraph', id: 'b', runs: [{ text: 'The team' }] },
				],
			},
		];
		const bytes = await saveDocx(model);
		const zip = await JSZip.loadAsync(bytes);
		const xml = await must(zip.file('word/glossary/document.xml')).async('string');
		expect(xml).toContain('<w:gallery w:val="docParts"/>');
		expect(xml).toContain('w:val="Sign-off"');
		expect(await must(zip.file('[Content_Types].xml')).async('string')).toContain(
			'document.glossary+xml',
		);
		expect(await must(zip.file('word/_rels/document.xml.rels')).async('string')).toContain(
			'glossaryDocument',
		);
		const loaded = await loadDocx(bytes);
		const entry = must(loaded.model.glossary?.[0]);
		expect(entry).toMatchObject({
			name: 'Sign-off',
			gallery: 'docParts',
			description: 'Closing line',
		});
		expect(entry.blocks).toHaveLength(2);
		expect(glossaryPreview(entry)).toBe('Kind regards,');
		// An untouched glossary is not rewritten.
		const same = await JSZip.loadAsync(await loaded.save(loaded.model));
		expect(await must(same.file('word/glossary/document.xml')).async('string')).toBe(xml);
	});

	it('adds to and deletes from an existing glossary, keeping the other entries', async () => {
		const first = createDocument();
		first.glossary = ['One', 'Two'].map((name) => ({
			name,
			gallery: 'autoTxt',
			category: 'General',
			blocks: [{ type: 'paragraph' as const, id: name, runs: [{ text: name }] }],
		}));
		const loaded = await loadDocx(await saveDocx(first));
		const model = loaded.model;
		model.glossary = [
			...must(model.glossary).filter((entry) => entry.name !== 'One'),
			{
				name: 'Three',
				gallery: 'docParts',
				category: 'General',
				blocks: [{ type: 'paragraph', id: 'c', runs: [{ text: 'Three' }] }],
			},
		];
		const xml = await must(
			(await JSZip.loadAsync(await loaded.save(model))).file('word/glossary/document.xml'),
		).async('string');
		expect(xml).not.toContain('w:val="One"');
		expect(xml).toContain('w:val="Two"');
		expect(xml).toContain('w:val="Three"');
		expect(glossaryEntryKey({ name: 'a', gallery: 'g', category: 'c' })).toContain('a');
	});

	it('replaces the entry of the same name, gallery and category when a selection is saved', async () => {
		const first = createDocument();
		first.glossary = [
			{
				name: 'Note',
				gallery: 'autoTxt',
				category: 'General',
				blocks: [{ type: 'paragraph', id: 'a', runs: [{ text: 'Old text' }] }],
			},
		];
		const loaded = await loadDocx(await saveDocx(first));
		loaded.model.glossary = [
			{
				name: 'Note',
				gallery: 'autoTxt',
				category: 'General',
				blocks: [{ type: 'paragraph', id: 'b', runs: [{ text: 'New text' }] }],
			},
		];
		const bytes = await loaded.save(loaded.model);
		const xml = await must((await JSZip.loadAsync(bytes)).file('word/glossary/document.xml')).async(
			'string',
		);
		expect(xml.match(/w:val="Note"/g)).toHaveLength(1);
		expect(xml).toContain('New text');
		expect(xml).not.toContain('Old text');
		const entries = must((await loadDocx(bytes)).model.glossary);
		expect(entries.map(glossaryPreview)).toEqual(['New text']);
	});

	it('keeps a glossary that Word wrote untouched and lists its entries', async () => {
		const zip = await JSZip.loadAsync(await saveDocx(createDocument()));
		const glossary = `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Footer note"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category></w:docPartPr><w:docPartBody><w:p><w:r><w:t>Footer text</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;
		zip.file('word/glossary/document.xml', glossary);
		zip.file(
			'word/_rels/document.xml.rels',
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		expect(glossaryPreview(must(loaded.model.glossary?.[0]))).toBe('Footer text');
		const saved = await JSZip.loadAsync(await loaded.save(loaded.model));
		expect(await must(saved.file('word/glossary/document.xml')).async('string')).toBe(glossary);
		expect(glossaryBlocksWritable(must(loaded.model.glossary?.[0]).blocks)).toBe(true);
	});
});
