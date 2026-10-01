import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	RELATIONSHIP_TYPES,
	buildContentTypesXml,
	buildRelationshipsXml,
	contentTypeForPart,
	ensureContentTypeDefault,
	ensureContentTypeOverride,
	ensureRelationship,
	findOfficeDocumentPart,
	getRelationshipId,
	isSafeHyperlinkHref,
	nextRelationshipId,
	parseContentTypes,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
} from './index.js';
import { NS, parseXml } from '@christophervr/ooxml-xml';

const RELS = `<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.styles}" Target="styles.xml"/><Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.hyperlink}" Target="https://a.test/?a=1&amp;b=2" TargetMode="External"/><Relationship Id="" Type="x" Target="skipped.xml"/></Relationships>`;

describe('relationships', () => {
	it('parses ids, types, targets and modes, skipping entries without an id', () => {
		const map = parseRelationships(RELS);
		expect([...map.keys()]).toEqual(['rId1', 'rId2']);
		expect(map.get('rId1')).toEqual({
			id: 'rId1',
			type: RELATIONSHIP_TYPES.styles,
			target: 'styles.xml',
			mode: 'Internal',
		});
		expect(map.get('rId2')).toMatchObject({ target: 'https://a.test/?a=1&b=2', mode: 'External' });
		expect(parseRelationships(undefined).size).toBe(0);
	});

	it('round-trips through buildRelationshipsXml with attribute escaping', () => {
		const rebuilt = buildRelationshipsXml(parseRelationships(RELS));
		expect(rebuilt).toContain('Target="https://a.test/?a=1&amp;b=2" TargetMode="External"');
		expect([...parseRelationships(rebuilt).values()]).toEqual([
			...parseRelationships(RELS).values(),
		]);
	});

	it('resolves targets against the declaring part', () => {
		expect(resolvePartPath('word/document.xml', 'media/image1.png')).toBe('word/media/image1.png');
		expect(resolvePartPath('word/document.xml', '../customXml/item1.xml')).toBe(
			'customXml/item1.xml',
		);
		expect(resolvePartPath('ppt/slides/slide1.xml', '../media/a.png')).toBe('ppt/media/a.png');
		expect(resolvePartPath('word/document.xml', '/docProps/core.xml')).toBe('docProps/core.xml');
		expect(resolvePartPath('word/document.xml', './header1.xml')).toBe('word/header1.xml');
		expect(relationshipsPartFor('word/document.xml')).toBe('word/_rels/document.xml.rels');
		expect(relationshipsPartFor('document.xml')).toBe('_rels/document.xml.rels');
	});

	it('allocates the first unused rId and reads r:id attributes', () => {
		expect(nextRelationshipId(new Set(['rId1', 'rId3']))).toBe('rId2');
		expect(nextRelationshipId(new Set())).toBe('rId1');
		const el = parseXml(
			`<w:headerReference xmlns:w="${NS.w}" xmlns:r="${NS.r}" r:id="rId7"/>`,
		).documentElement;
		expect(getRelationshipId(el)).toBe('rId7');
		expect(getRelationshipId(parseXml('<a/>').documentElement)).toBeUndefined();
	});
});

describe('content types', () => {
	const XML = `<Types xmlns="${NS.contentTypes}"><Default Extension="PNG" ContentType="image/png"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="main+xml"/></Types>`;

	it('parses defaults (extension lower-cased) and overrides and resolves a part type', () => {
		const types = parseContentTypes(XML);
		expect(types.defaults.get('png')).toBe('image/png');
		expect(contentTypeForPart(types, 'word/document.xml')).toBe('main+xml');
		expect(contentTypeForPart(types, '/word/media/a.PNG')).toBe('image/png');
		expect(contentTypeForPart(types, 'word/other.bin')).toBeUndefined();
		expect(parseContentTypes(undefined).overrides.size).toBe(0);
	});

	it('adds a default only when absent and rebuilds valid XML', () => {
		expect(ensureContentTypeDefault(XML, 'png', 'image/png')).toBe(XML);
		const added = ensureContentTypeDefault(XML, 'gif', 'image/gif');
		expect(parseContentTypes(added).defaults.get('gif')).toBe('image/gif');
		expect(parseContentTypes(buildContentTypesXml(parseContentTypes(XML))).overrides.size).toBe(1);
	});
});

describe('zip-level helpers', () => {
	it('adds a content-type override once and creates the part when missing', async () => {
		const zip = new JSZip();
		await ensureContentTypeOverride(zip, 'word/numbering.xml', 'numbering+xml');
		await ensureContentTypeOverride(zip, 'word/numbering.xml', 'numbering+xml');
		const xml = await zip.file('[Content_Types].xml')!.async('string');
		expect(xml.match(/<Override /g)).toHaveLength(1);
		expect(parseContentTypes(xml).overrides.get('/word/numbering.xml')).toBe('numbering+xml');
	});

	it('reuses an existing relationship and numbers new ones after the highest id', async () => {
		const zip = new JSZip();
		zip.file('word/_rels/document.xml.rels', RELS);
		expect(
			await ensureRelationship(
				zip,
				'word/_rels/document.xml.rels',
				RELATIONSHIP_TYPES.styles,
				'styles.xml',
			),
		).toBe('rId1');
		const id = await ensureRelationship(
			zip,
			'word/_rels/document.xml.rels',
			RELATIONSHIP_TYPES.numbering,
			'numbering.xml',
		);
		expect(id).toBe('rId3');
		const created = await ensureRelationship(
			new JSZip(),
			'x/_rels/y.rels',
			RELATIONSHIP_TYPES.theme,
			't.xml',
		);
		expect(created).toBe('rId1');
		const xml = await zip.file('word/_rels/document.xml.rels')!.async('string');
		expect(xml).toContain('Target="numbering.xml"');
		expect(xml).toContain('skipped.xml');
	});

	it('finds the main document part from the package relationships', async () => {
		const zip = new JSZip();
		expect(await findOfficeDocumentPart(zip)).toBeUndefined();
		zip.file(
			'_rels/.rels',
			`<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
		);
		expect(await findOfficeDocumentPart(zip)).toBe('word/document.xml');
	});
});

describe('isSafeHyperlinkHref', () => {
	it('accepts only http, https and mailto', () => {
		for (const ok of ['http://a', 'HTTPS://a', ' mailto:x@y.z'])
			expect(isSafeHyperlinkHref(ok)).toBe(true);
		for (const bad of ['javascript:alert(1)', 'data:text/html,x', '/relative', 'file:///c', ''])
			expect(isSafeHyperlinkHref(bad)).toBe(false);
	});
});
