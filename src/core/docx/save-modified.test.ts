import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { expectParagraph, must } from './test-support/access';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CORE = `<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Old</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created></cp:coreProperties>`;

async function fixture(withCore: boolean): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	if (withCore) zip.file('docProps/core.xml', CORE);
	return zip.generateAsync({ type: 'uint8array' });
}

const at = new Date('2026-10-29T10:11:12.345Z');
const coreOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('docProps/core.xml')?.async('string')) ?? '';
async function edited(withCore = true) {
	const loaded = await loadDocx(await fixture(withCore));
	must(expectParagraph(loaded.model.blocks[0]).runs[0]).text = 'Edited';
	return loaded;
}

describe('last-modified stamp on save', () => {
	it('records the save time and the host author, keeping other elements', async () => {
		const loaded = await edited();
		const xml = await coreOf(
			await loaded.save(loaded.model, undefined, { modified: { at, author: 'Bob' } }),
		);
		expect(xml).toContain('>2026-10-29T10:11:12Z</dcterms:modified>');
		expect(xml).toContain('xsi:type="dcterms:W3CDTF"');
		expect(xml).toContain('<cp:lastModifiedBy>Bob</cp:lastModifiedBy>');
		expect(xml).toContain('<dc:title>Old</dc:title>');
		expect(xml).toContain('2024-01-02T03:04:05Z');
	});

	it('is off unless asked for and stamps the current time when given', async () => {
		const loaded = await edited();
		expect(await coreOf(await loaded.save(loaded.model))).not.toContain('dcterms:modified');
		const stamped = await coreOf(await loaded.save(loaded.model, undefined, { modified: {} }));
		expect(stamped).toMatch(/<dcterms:modified[^>]*>20\d\d-/);
	});

	it('leaves an unchanged document and a package without core properties alone', async () => {
		const bytes = await fixture(true);
		const same = await (await loadDocx(bytes)).save(undefined, undefined, { modified: { at } });
		expect(same).toEqual(bytes);
		const loaded = await edited(false);
		expect(await coreOf(await loaded.save(loaded.model, undefined, { modified: { at } }))).toBe('');
	});
});
