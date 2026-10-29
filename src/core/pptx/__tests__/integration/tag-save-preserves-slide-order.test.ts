import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';

/**
 * Regression for issue #232: freeform icons came out of the editor as tangles.
 *
 * Every binding saves with the deck's loaded tag collections (`tags`), and the
 * tag pass re-parsed each already-written slide that owns tags and rebuilt it
 * through the grouping object model. That regrouped every interleaved custom
 * geometry path (all `moveTo`s, then all `lnTo`s...), and re-placed each
 * shape's `p:tags` onto the slide, so the slide ended up pointing at the last
 * shape's tags part. A plain `save(slides)` never ran the pass, which is why
 * only the editors showed it.
 */
const TAGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags';
const TAGS_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml';

const FREEFORM =
	'<p:sp><p:nvSpPr><p:cNvPr id="90" name="Icon"/><p:cNvSpPr/><p:nvPr><p:custDataLst><p:tags r:id="rIdTagShape"/></p:custDataLst></p:nvPr></p:nvSpPr>' +
	'<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100" h="100">' +
	'<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
	'<a:cubicBezTo><a:pt x="100" y="50"/><a:pt x="50" y="100"/><a:pt x="0" y="100"/></a:cubicBezTo>' +
	'<a:lnTo><a:pt x="-4" y="50"/></a:lnTo><a:close/>' +
	'<a:moveTo><a:pt x="20" y="20"/></a:moveTo><a:lnTo><a:pt x="40" y="20"/></a:lnTo><a:close/>' +
	'</a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:spPr></p:sp>';

function tagPart(name: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag name="${name}" val="1"/></p:tagLst>`;
}

/** A one-slide deck whose freeform owns a tags part, and whose slide owns another. */
async function taggedFreeformDeck(): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(createSlide('Blank').build());
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
	zip.file(
		'ppt/slides/slide1.xml',
		slide.replace(
			/<\/p:spTree>/u,
			`${FREEFORM}</p:spTree><p:custDataLst><p:tags r:id="rIdTagSlide"/></p:custDataLst>`,
		),
	);
	const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
	zip.file(
		'ppt/slides/_rels/slide1.xml.rels',
		rels.replace(
			'</Relationships>',
			`<Relationship Id="rIdTagShape" Type="${TAGS_REL}" Target="../tags/tag1.xml"/>` +
				`<Relationship Id="rIdTagSlide" Type="${TAGS_REL}" Target="../tags/tag2.xml"/></Relationships>`,
		),
	);
	zip.file('ppt/tags/tag1.xml', tagPart('SHAPE'));
	zip.file('ppt/tags/tag2.xml', tagPart('SLIDE'));
	const types = await zip.file('[Content_Types].xml')!.async('string');
	zip.file(
		'[Content_Types].xml',
		types.replace(
			'</Types>',
			`<Override PartName="/ppt/tags/tag1.xml" ContentType="${TAGS_TYPE}"/>` +
				`<Override PartName="/ppt/tags/tag2.xml" ContentType="${TAGS_TYPE}"/></Types>`,
		),
	);
	return zip.generateAsync({ type: 'uint8array' });
}

async function slideXml(bytes: Uint8Array): Promise<string> {
	return (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml')!.async('string');
}

const EXPECTED_COMMANDS = [
	'moveTo',
	'lnTo',
	'cubicBezTo',
	'lnTo',
	'close',
	'moveTo',
	'lnTo',
	'close',
];

function commandSequence(xml: string): string[] {
	const path = /<a:path\b[^>]*>([\s\S]*?)<\/a:path>/u.exec(xml)?.[1] ?? '';
	return [...path.matchAll(/<a:(moveTo|lnTo|cubicBezTo|close)\b/gu)].map((match) => match[1]!);
}

describe('saving a deck with its tag collections (issue #232)', () => {
	it('keeps custom-geometry command order and per-shape tag ownership', async () => {
		const source = await taggedFreeformDeck();
		const handler = new PptxHandler();
		const data = await handler.load(source.buffer as ArrayBuffer);
		expect(data.tags?.length).toBe(2);

		// The editors pass the slides back as new objects, so the slide is rebuilt.
		const saved = await handler.save(
			data.slides.map((slide) => ({ ...slide })),
			{ tags: data.tags },
		);
		const xml = await slideXml(saved);

		expect(commandSequence(xml)).toEqual(EXPECTED_COMMANDS);
		expect(xml).toMatch(/<p:nvPr><p:custDataLst><p:tags r:id="rIdTagShape"/u);
		expect(xml).toMatch(/<\/p:spTree><p:custDataLst><p:tags r:id="rIdTagSlide"/u);
		expect(xml.match(/<p:tags /gu)).toHaveLength(2);
	});

	it('keeps them when the slide also authors customer data', async () => {
		const source = await taggedFreeformDeck();
		const handler = new PptxHandler();
		const data = await handler.load(source.buffer as ArrayBuffer);
		const slides = data.slides.map((slide) => ({ ...slide }));
		slides[0]!.customerData = [{ data: '<slide-data/>' }];

		const xml = await slideXml(await handler.save(slides, { tags: data.tags }));

		expect(commandSequence(xml)).toEqual(EXPECTED_COMMANDS);
		expect(xml).toMatch(/<p:nvPr><p:custDataLst><p:tags r:id="rIdTagShape"/u);
		// CT_CustomerDataList is `custData*, tags?`: the entry goes ahead of the slide's tags.
		expect(xml).toMatch(
			/<\/p:spTree><p:custDataLst><p:custData r:id="rId\d+"\/><p:tags r:id="rIdTagSlide"\/><\/p:custDataLst>/u,
		);
	});
});
