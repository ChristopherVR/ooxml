import { DOMParser } from '@xmldom/xmldom';
import { describe, it, expect } from 'vitest';

import { applyReferenceTransforms, extractReferenceTransforms } from './reference-transforms';
import type { XmlDocument } from './xml-canonicalization';

/** Parses with xmldom, typed as the structural document the modules accept (as their callers do). */
const parseXml = (xml: string): XmlDocument =>
	new DOMParser().parseFromString(xml, 'text/xml') as unknown as XmlDocument;

// ---------------------------------------------------------------------------
// extractReferenceTransforms
// ---------------------------------------------------------------------------

describe('extractReferenceTransforms', () => {
	it('parses transforms with correct algorithm from a Reference element', () => {
		const xml = [
			'<ds:Reference URI="/part.xml" xmlns:ds="http://www.w3.org/2000/09/xmldsig#">',
			'  <ds:Transforms>',
			'    <ds:Transform Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"/>',
			'  </ds:Transforms>',
			'</ds:Reference>',
		].join('\n');
		const doc = parseXml(xml);
		const referenceElement = doc.documentElement!;
		const transforms = extractReferenceTransforms(referenceElement);

		expect(transforms).toHaveLength(1);
		expect(transforms[0]?.algorithm).toBe('http://www.w3.org/2001/10/xml-exc-c14n#');
		expect(transforms[0]?.relationshipReferenceIds).toStrictEqual([]);
	});

	it('parses multiple transforms', () => {
		const xml = [
			'<ds:Reference URI="/part.xml" xmlns:ds="http://www.w3.org/2000/09/xmldsig#">',
			'  <ds:Transforms>',
			'    <ds:Transform Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"/>',
			'    <ds:Transform Algorithm="http://www.w3.org/2000/09/xmldsig#enveloped-signature"/>',
			'  </ds:Transforms>',
			'</ds:Reference>',
		].join('\n');
		const doc = parseXml(xml);
		const transforms = extractReferenceTransforms(doc.documentElement!);

		expect(transforms).toHaveLength(2);
		expect(transforms[0]?.algorithm).toBe('http://www.w3.org/2001/10/xml-exc-c14n#');
		expect(transforms[1]?.algorithm).toBe('http://www.w3.org/2000/09/xmldsig#enveloped-signature');
	});

	it('returns empty array when no transforms present', () => {
		const xml = '<ds:Reference URI="/part.xml" xmlns:ds="http://www.w3.org/2000/09/xmldsig#"/>';
		const doc = parseXml(xml);
		const transforms = extractReferenceTransforms(doc.documentElement!);

		expect(transforms).toStrictEqual([]);
	});

	it('extracts RelationshipReference SourceId values', () => {
		const xml = [
			'<ds:Reference URI="/part.xml" xmlns:ds="http://www.w3.org/2000/09/xmldsig#">',
			'  <ds:Transforms>',
			'    <ds:Transform Algorithm="http://schemas.openxmlformats.org/package/2006/RelationshipTransform">',
			'      <mdssi:RelationshipReference xmlns:mdssi="http://schemas.openxmlformats.org/package/2006/digital-signature" SourceId="rId1"/>',
			'      <mdssi:RelationshipReference xmlns:mdssi="http://schemas.openxmlformats.org/package/2006/digital-signature" SourceId="rId2"/>',
			'    </ds:Transform>',
			'  </ds:Transforms>',
			'</ds:Reference>',
		].join('\n');
		const doc = parseXml(xml);
		const transforms = extractReferenceTransforms(doc.documentElement!);

		expect(transforms).toHaveLength(1);
		expect(transforms[0]?.algorithm).toBe(
			'http://schemas.openxmlformats.org/package/2006/RelationshipTransform',
		);
		expect(transforms[0]?.relationshipReferenceIds).toStrictEqual(['rId1', 'rId2']);
	});

	it('extracts RelationshipsGroupReference SourceType values', () => {
		const xml = [
			'<ds:Reference URI="/part.xml" xmlns:ds="http://www.w3.org/2000/09/xmldsig#">',
			'  <ds:Transforms>',
			'    <ds:Transform Algorithm="http://schemas.openxmlformats.org/package/2006/RelationshipTransform">',
			'      <mdssi:RelationshipsGroupReference xmlns:mdssi="http://schemas.openxmlformats.org/package/2006/digital-signature" SourceType="urn:type"/>',
			'    </ds:Transform>',
			'  </ds:Transforms>',
			'</ds:Reference>',
		].join('\n');
		const transforms = extractReferenceTransforms(parseXml(xml).documentElement!);

		expect(transforms[0]?.relationshipGroupTypes).toStrictEqual(['urn:type']);
	});
});

describe('applyReferenceTransforms', () => {
	const RELATIONSHIP_TRANSFORM =
		'http://schemas.openxmlformats.org/package/2006/RelationshipTransform';
	const C14N = 'http://www.w3.org/TR/2001/REC-xml-c14n-20010315';
	const NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
	const part = new TextEncoder().encode(
		`<?xml version="1.0"?>\n<r:Relationships xmlns:r="${NS}">\n` +
			'  <r:Relationship Type="urn:b" Target="b.xml" Id="rId2"/>\n' +
			'  <!-- note -->\n' +
			'  <r:Relationship Target="https://example.test/?a=1&amp;b=2" Id="rId10" Type="urn:c" TargetMode="External"/>\n' +
			'  <r:Relationship Id="rId1" Target="a.xml" Type="urn:a"/>\n' +
			'</r:Relationships>',
	);
	const relationshipTransform = (ids: string[], types: string[] = []) => ({
		algorithm: RELATIONSHIP_TRANSFORM,
		relationshipReferenceIds: ids,
		relationshipGroupTypes: types,
	});
	const c14n = { algorithm: C14N, relationshipReferenceIds: [], relationshipGroupTypes: [] };
	const text = (data: Uint8Array) => new TextDecoder().decode(data);

	it('produces the sorted, normalized form ECMA-376 Part 2 defines', () => {
		const result = applyReferenceTransforms(part, [
			relationshipTransform(['rId2', 'rId10'], ['urn:a']),
			c14n,
		]);

		expect(text(result.data)).toBe(
			`<Relationships xmlns="${NS}">` +
				'<Relationship Id="rId1" Target="a.xml" TargetMode="Internal" Type="urn:a"></Relationship>' +
				'<Relationship Id="rId10" Target="https://example.test/?a=1&amp;b=2" TargetMode="External" Type="urn:c"></Relationship>' +
				'<Relationship Id="rId2" Target="b.xml" TargetMode="Internal" Type="urn:b"></Relationship>' +
				'</Relationships>',
		);
		expect(result.selectedRelationships?.map((relationship) => relationship.id)).toStrictEqual([
			'rId1',
			'rId10',
			'rId2',
		]);
	});

	it('selects nothing when the transform names nothing', () => {
		const result = applyReferenceTransforms(part, [relationshipTransform([]), c14n]);

		expect(text(result.data)).toBe(`<Relationships xmlns="${NS}"></Relationships>`);
	});

	it('reports a transform it cannot apply', () => {
		const algorithm = 'http://www.w3.org/2000/09/xmldsig#enveloped-signature';
		const result = applyReferenceTransforms(part, [
			{ algorithm, relationshipReferenceIds: [], relationshipGroupTypes: [] },
		]);

		expect(result.unsupportedAlgorithms).toStrictEqual([algorithm]);
	});

	it('reports malformed relationships instead of digesting them', () => {
		const result = applyReferenceTransforms(new TextEncoder().encode('<Relationships'), [
			relationshipTransform(['rId1']),
		]);

		expect(result.unsupportedAlgorithms).toStrictEqual([RELATIONSHIP_TRANSFORM]);
	});
});
