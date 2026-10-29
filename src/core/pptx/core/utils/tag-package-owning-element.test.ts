import JSZip from 'jszip';
import { XMLBuilder, XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';

import type { PptxTagCollection } from '../types';
import { writeTagCollections } from './tag-package';

const SLIDE_PATH = 'ppt/slides/slide1.xml';
const RELS_PATH = 'ppt/slides/_rels/slide1.xml.rels';
const TAGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags';

/** A freeform whose path interleaves commands, owned tags on a shape AND on the slide. */
const SLIDE_XML =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
	'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
	'<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
	'<p:sp><p:nvSpPr><p:cNvPr id="2" name="Freeform"/><p:cNvSpPr/><p:nvPr><p:custDataLst><p:tags r:id="rId1"/></p:custDataLst></p:nvPr></p:nvSpPr>' +
	'<p:spPr><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100" h="100">' +
	'<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
	'<a:cubicBezTo><a:pt x="100" y="50"/><a:pt x="50" y="100"/><a:pt x="0" y="100"/></a:cubicBezTo>' +
	'<a:lnTo><a:pt x="-4" y="50"/></a:lnTo><a:close/>' +
	'<a:moveTo><a:pt x="20" y="20"/></a:moveTo><a:lnTo><a:pt x="40" y="20"/></a:lnTo><a:close/>' +
	'</a:path></a:pathLst></a:custGeom></p:spPr></p:sp>' +
	'</p:spTree><p:custDataLst><p:tags r:id="rId2"/></p:custDataLst></p:cSld>' +
	'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';

const RELS_XML =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
	'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
	`<Relationship Id="rId2" Type="${TAGS_REL}" Target="../tags/tag2.xml"/>` +
	`<Relationship Id="rId1" Type="${TAGS_REL}" Target="../tags/tag1.xml"/>` +
	'</Relationships>';

/** The save pipeline's own codec shape: a grouping parser and a plain builder. */
const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_' });
const codec = {
	parse: (xml: string) => parser.parse(xml) as Record<string, unknown>,
	build: (data: Record<string, unknown>) => builder.build(data) as string,
};

function deck(): JSZip {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
	);
	zip.file(SLIDE_PATH, SLIDE_XML);
	zip.file(RELS_PATH, RELS_XML);
	return zip;
}

function collection(index: number, relationshipId: string): PptxTagCollection {
	return {
		path: `ppt/tags/tag${index}.xml`,
		owner: 'slide',
		sourcePartPath: SLIDE_PATH,
		relationshipId,
		tags: [{ name: `NAME${index}`, value: `value ${index}` }],
	};
}

/** The path commands of the first `a:path`, in document order. */
function commandSequence(xml: string): string[] {
	const path = /<a:path\b[^>]*>([\s\S]*?)<\/a:path>/u.exec(xml)?.[1] ?? '';
	return [...path.matchAll(/<a:(moveTo|lnTo|cubicBezTo|close)\b/gu)].map((match) => match[1]!);
}

describe('writeTagCollections owning element', () => {
	it('leaves a slide whose shape and slide tags are already owned byte-for-byte unchanged', async () => {
		const zip = deck();
		await writeTagCollections(zip, [collection(1, 'rId1'), collection(2, 'rId2')], codec);
		expect(await zip.file(SLIDE_PATH)!.async('string')).toBe(SLIDE_XML);
	});

	it('adds a new slide-owned collection without regrouping custom-geometry commands', async () => {
		const zip = deck();
		const added: PptxTagCollection = {
			owner: 'slide',
			sourcePartPath: SLIDE_PATH,
			tags: [{ name: 'NEW', value: 'yes' }],
		};
		await writeTagCollections(zip, [added], codec);
		const xml = await zip.file(SLIDE_PATH)!.async('string');
		expect(commandSequence(xml)).toEqual(commandSequence(SLIDE_XML));
		expect(commandSequence(xml)).toEqual([
			'moveTo',
			'lnTo',
			'cubicBezTo',
			'lnTo',
			'close',
			'moveTo',
			'lnTo',
			'close',
		]);
		// The shape keeps its own tags; the slide's single list now owns the new part.
		expect(xml).toContain('<p:nvPr><p:custDataLst><p:tags r:id="rId1"/></p:custDataLst></p:nvPr>');
		expect(xml).toContain(
			`</p:spTree><p:custDataLst><p:tags r:id="${added.relationshipId}"/></p:custDataLst></p:cSld>`,
		);
		expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')).toBe(
			true,
		);
	});

	it('removes only the emptied collection, wherever its owning element sits', async () => {
		const zip = deck();
		await writeTagCollections(zip, [{ ...collection(1, 'rId1'), tags: [] }], codec);
		const xml = await zip.file(SLIDE_PATH)!.async('string');
		expect(xml).not.toContain('r:id="rId1"');
		expect(xml).toContain('<p:nvPr/>');
		expect(xml).toContain('<p:custDataLst><p:tags r:id="rId2"/></p:custDataLst></p:cSld>');
		expect(commandSequence(xml)).toEqual(commandSequence(SLIDE_XML));
	});
});
