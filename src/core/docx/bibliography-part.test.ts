import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { BIBLIOGRAPHY_NS, readBibliographyPart } from './bibliography-part';

const zipWith = (parts: Record<string, string>): JSZip => {
	const zip = new JSZip();
	for (const [name, xml] of Object.entries(parts)) zip.file(name, xml);
	return zip;
};

describe('readBibliographyPart', () => {
	it('finds the part whose root is a bibliography Sources element, whatever its prefix', async () => {
		const sources = `<b:Sources xmlns:b="${BIBLIOGRAPHY_NS}" SelectedStyle="APA"/>`;
		const zip = zipWith({
			'customXml/item1.xml': '<root xmlns="urn:other"/>',
			'customXml/item2.xml': sources,
		});
		expect(await readBibliographyPart(zip, ['customXml/item1.xml', 'customXml/item2.xml'])).toEqual(
			{ part: 'customXml/item2.xml', xml: sources },
		);
	});

	it('reads the namespace from the parsed root, not from text that mentions it', async () => {
		const zip = zipWith({
			// The namespace as text, on a host whose name merely contains it, or on another root.
			'customXml/item1.xml': `<Sources xmlns="urn:other"><note>${BIBLIOGRAPHY_NS}</note></Sources>`,
			'customXml/item2.xml': `<Sources xmlns="https://a.test/${BIBLIOGRAPHY_NS}"/>`,
			'customXml/item3.xml': `<b:Source xmlns:b="${BIBLIOGRAPHY_NS}"/>`,
		});
		expect(
			await readBibliographyPart(zip, [
				'customXml/item1.xml',
				'customXml/item2.xml',
				'customXml/item3.xml',
			]),
		).toBeUndefined();
	});
});
