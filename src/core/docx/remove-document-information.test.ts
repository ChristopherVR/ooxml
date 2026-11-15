import { describe, expect, it } from 'vitest';
import { loadDocx } from './parse';
import {
	inspectDocument,
	INSPECTION_CATEGORIES,
	type InspectionCategory,
} from './inspect-document';
import { removeDocumentInformation } from './remove-document-information';
import { AUTHOR, inspectFixture, REVIEWER, xmlParts } from './test-support/inspect-fixture';

/** Loads the fixture, removes `categories`, saves (as the editor does) and reloads. */
async function removeAndSave(categories: readonly InspectionCategory[]) {
	const loaded = await loadDocx(await inspectFixture());
	const next = removeDocumentInformation(loaded.model, categories);
	const bytes = await loaded.save(next, undefined, {
		modified: { at: new Date('2026-02-01T00:00:00Z'), author: REVIEWER },
	});
	const reloaded = await loadDocx(bytes);
	return { next, parts: await xmlParts(bytes), reloaded: reloaded.model };
}

const findings = (model: Parameters<typeof inspectDocument>[0], category: InspectionCategory) =>
	inspectDocument(model, [category])[0]!.findings;

describe('removeDocumentInformation', () => {
	it('removes comments with their parts and accepts every tracked change', async () => {
		const { parts, reloaded } = await removeAndSave(['comments']);
		expect(findings(reloaded, 'comments')).toEqual([]);
		const document = parts.get('word/document.xml')!;
		expect(document).not.toMatch(/commentRange|commentReference|<w:ins\b/);
		expect(document).toContain('Tracked insertion.');
		// Word deletes the emptied comment parts; they are kept here, holding no comments.
		expect(parts.get('word/comments.xml')).not.toContain('<w:comment ');
	});

	it('removes properties and personal information as Word does, with no names left', async () => {
		const { parts, reloaded } = await removeAndSave(['properties']);
		expect(findings(reloaded, 'properties')).toEqual([]);
		for (const [name, xml] of parts) {
			expect(xml, name).not.toContain(AUTHOR);
			expect(xml, name).not.toContain(REVIEWER);
			expect(xml, name).not.toMatch(/\sw:date=|dateUtc=/);
		}
		const document = parts.get('word/document.xml')!;
		expect(document).toContain('w:author="Author"');
		expect(parts.get('word/comments.xml')).toMatch(/w:author="Author"[^>]*w:initials="A"/);
		const settings = parts.get('word/settings.xml')!;
		expect(settings).toMatch(
			/<w:zoom[^>]*\/><w:removePersonalInformation\/><w:removeDateAndTime\/>/,
		);
		expect(settings).not.toMatch(/attachedTemplate|w15:docId/);
		expect(parts.has('word/_rels/settings.xml.rels')).toBe(false);
		expect(parts.has('word/people.xml')).toBe(false);
		expect(parts.has('word/commentsExtensible.xml')).toBe(false);
		expect(parts.has('docProps/custom.xml')).toBe(false);
		expect(parts.get('_rels/.rels')).not.toContain('custom.xml');
		const core = parts.get('docProps/core.xml')!;
		expect(core).not.toMatch(/dc:title|cp:keywords|cp:category|dc:creator|cp:lastModifiedBy/);
		// Word keeps the dates and revision number of the package itself.
		expect(core).toMatch(/dcterms:created[^>]*>2026-01-02T03:04:00Z/);
		const app = parts.get('docProps/app.xml')!;
		expect(app).not.toMatch(/Company|Manager/);
		expect(app).toContain('<Template>Normal.dotm</Template>');
		expect(reloaded.packageInformation).toMatchObject({
			removePersonalInformation: true,
			removeDateAndTime: true,
			customXmlParts: ['customXml/item1.xml'],
		});
	});

	it('keeps personal information out of later saves while the setting is on', async () => {
		const loaded = await loadDocx(await inspectFixture());
		const removed = await loaded.save(removeDocumentInformation(loaded.model, ['properties']));
		const again = await loadDocx(removed);
		const later = {
			id: '9',
			author: REVIEWER,
			initials: 'BR',
			date: '2026-03-01T00:00:00Z',
			text: 'Later',
		};
		const bytes = await again.save(
			{ ...again.model, comments: [...(again.model.comments ?? []), later] },
			undefined,
			{ modified: { author: REVIEWER } },
		);
		for (const [name, xml] of await xmlParts(bytes)) {
			expect(xml, name).not.toContain(REVIEWER);
			expect(xml, name).not.toMatch(/\sw:date=|dateUtc=/);
		}
	});

	it('removes custom XML data parts and their relationships', async () => {
		const { parts, reloaded } = await removeAndSave(['customXml']);
		expect(findings(reloaded, 'customXml')).toEqual([]);
		expect([...parts.keys()].some((name) => name.startsWith('customXml/'))).toBe(false);
		expect(parts.get('word/_rels/document.xml.rels')).not.toContain('customXml');
		expect(parts.get('[Content_Types].xml')).not.toContain('itemProps1');
		// Other categories are left alone.
		expect(parts.get('word/people.xml')).toContain(AUTHOR);
	});

	it('empties headers and footers but keeps their parts and formatting', async () => {
		const { parts, reloaded } = await removeAndSave(['headersFooters']);
		expect(findings(reloaded, 'headersFooters')).toEqual([]);
		expect(parts.get('word/header1.xml')).toMatch(/<w:pStyle w:val="Header"\/>/);
		expect(parts.get('word/header1.xml')).not.toContain('Header text');
		expect(parts.get('word/footer1.xml')).not.toContain('Footer text');
		expect(parts.get('word/document.xml')).toMatch(/headerReference[^>]*default/);
	});

	it('deletes hidden text and keeps its paragraph', async () => {
		const { parts, reloaded } = await removeAndSave(['hiddenText']);
		expect(findings(reloaded, 'hiddenText')).toEqual([]);
		const document = parts.get('word/document.xml')!;
		expect(document).not.toMatch(/hidden words|w:vanish/);
		expect(document).toContain('Shown');
	});

	it('removes everything, leaving nothing to report', async () => {
		const { reloaded } = await removeAndSave(INSPECTION_CATEGORIES);
		expect(inspectDocument(reloaded).every((entry) => !entry.findings.length)).toBe(true);
	});

	it('leaves the model unchanged when a category has nothing to remove', async () => {
		const { model } = await loadDocx(await inspectFixture());
		const clean = removeDocumentInformation(model, ['customXml', 'hiddenText', 'headersFooters']);
		expect(removeDocumentInformation(clean, ['customXml'])).toBe(clean);
	});
});
