import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { createDocument } from './model';
import {
	CORE_PROPERTIES_STORE,
	DOCUMENT_PROPERTY_CHOICES,
	documentPropertyControl,
	documentPropertyValue,
	propertyBinding,
	documentPropertyChoice,
} from './quick-parts';
import { saveDocx } from './save';
import { expectParagraph, must } from './test-support/access';

const CORE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Ada</dc:creator></cp:coreProperties>`;
const APP = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Test</Application></Properties>`;

async function withProperties(): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await saveDocx(createDocument()));
	zip.file('docProps/core.xml', CORE);
	zip.file('docProps/app.xml', APP);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('Quick Parts document properties', () => {
	it('binds Author to the core properties part the way Word does', () => {
		const choice = must(documentPropertyChoice('author'));
		expect(propertyBinding(choice)).toEqual({
			prefixMappings:
				"xmlns:ns0='http://purl.org/dc/elements/1.1/' xmlns:ns1='http://schemas.openxmlformats.org/package/2006/metadata/core-properties'",
			xpath: '/ns1:coreProperties[1]/ns0:creator[1]',
			storeItemId: CORE_PROPERTIES_STORE,
		});
		expect(DOCUMENT_PROPERTY_CHOICES.map((entry) => entry.id)).toContain('company');
	});

	it('saves a data-bound control that reloads as bound and writes edits to the property part', async () => {
		const loaded = await loadDocx(await withProperties());
		expect(documentPropertyValue(loaded.model, 'author')).toBe('Ada');
		const { control } = must(documentPropertyControl('author', new Set()));
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.runs = [{ text: 'Grace', contentControl: control }];
		const saved = await loaded.save(loaded.model);
		const zip = await JSZip.loadAsync(saved);
		const xml = await must(zip.file('word/document.xml')).async('string');
		expect(xml).toContain('w:storeItemID="{6C3C8BC8-F283-45AE-878A-BAB7291924A1}"');
		expect(xml).toContain('/ns1:coreProperties[1]/ns0:creator[1]');
		expect(xml).toContain('<w:alias w:val="Author"/>');
		expect(await must(zip.file('docProps/core.xml')).async('string')).toContain(
			'<dc:creator>Grace</dc:creator>',
		);
		const again = await loadDocx(saved);
		const run = must(expectParagraph(again.model.blocks[0]).runs[0]);
		expect(run.contentControl?.dataBound).toBe(true);
		expect(run.contentControl?.alias).toBe('Author');
	});

	it('adds a missing extended property element such as Company', async () => {
		const loaded = await loadDocx(await withProperties());
		const { control } = must(documentPropertyControl('company', new Set()));
		expectParagraph(loaded.model.blocks[0]).runs = [{ text: 'Contoso', contentControl: control }];
		const zip = await JSZip.loadAsync(await loaded.save(loaded.model));
		const app = await must(zip.file('docProps/app.xml')).async('string');
		expect(app).toContain('Company>Contoso</');
		expect(app).toContain('<Application>Test</Application>');
	});
});
