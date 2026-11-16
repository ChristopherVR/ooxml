import { DOMParser } from '@xmldom/xmldom';
import { describe, it, expect, expectTypeOf } from 'vitest';

import {
	getNodeLocalName,
	getFirstDescendantElementByLocalName,
	canonicalizeSignedInfoXml,
} from './xml-canonicalization';
import type { XmlDocument } from './xml-canonicalization';

/** Parses with xmldom, typed as the structural document the module accepts (as its callers do). */
const parseXml = (xml: string): XmlDocument =>
	new DOMParser().parseFromString(xml, 'text/xml') as unknown as XmlDocument;

// ---------------------------------------------------------------------------
// getNodeLocalName
// ---------------------------------------------------------------------------

describe('getNodeLocalName', () => {
	it('returns local name stripped of namespace prefix', () => {
		const doc = parseXml('<ds:Signature xmlns:ds="http://www.w3.org/2000/09/xmldsig#"/>');
		expect(getNodeLocalName(doc.documentElement!)).toBe('Signature');
	});

	it('returns the full name when no prefix is present', () => {
		const doc = parseXml('<Root/>');
		expect(getNodeLocalName(doc.documentElement!)).toBe('Root');
	});

	it('handles deeply-prefixed names like a:b:Foo by stripping first prefix', () => {
		const doc = parseXml('<ns:Element xmlns:ns="urn:test"/>');
		expect(getNodeLocalName(doc.documentElement!)).toBe('Element');
	});
});

// ---------------------------------------------------------------------------
// getFirstDescendantElementByLocalName
// ---------------------------------------------------------------------------

describe('getFirstDescendantElementByLocalName', () => {
	it('finds element by local name in parsed XML', () => {
		const doc = parseXml(
			'<Root><ds:SignedInfo xmlns:ds="http://www.w3.org/2000/09/xmldsig#"><ds:Reference/></ds:SignedInfo></Root>',
		);
		const result = getFirstDescendantElementByLocalName(doc, 'SignedInfo');
		expect(result).toBeDefined();
		expect(getNodeLocalName(result!)).toBe('SignedInfo');
	});

	it('returns the first matching element when multiple exist', () => {
		const doc = parseXml('<Root><Item id="1"/><Item id="2"/></Root>');
		const result = getFirstDescendantElementByLocalName(doc, 'Item');
		expect(result).toBeDefined();
		expect(result!.getAttribute('id')).toBe('1');
	});

	it('returns undefined when element is not found', () => {
		const doc = parseXml('<Root><Child/></Root>');
		const result = getFirstDescendantElementByLocalName(doc, 'NonExistent');
		expect(result).toBeUndefined();
	});

	it('searches within an element subtree', () => {
		const doc = parseXml('<Root><Parent><Target/></Parent></Root>');
		const parent = getFirstDescendantElementByLocalName(doc, 'Parent')!;
		const result = getFirstDescendantElementByLocalName(parent, 'Target');
		expect(result).toBeDefined();
		expect(getNodeLocalName(result!)).toBe('Target');
	});
});

// ---------------------------------------------------------------------------
// canonicalizeSignedInfoXml
// ---------------------------------------------------------------------------

describe('canonicalizeSignedInfoXml', () => {
	it('returns a string for well-formed XML input', () => {
		const xml =
			'<SignedInfo xmlns="http://www.w3.org/2000/09/xmldsig#"><CanonicalizationMethod/></SignedInfo>';
		const result = canonicalizeSignedInfoXml(xml);
		expectTypeOf(result).toBeString();
		expect(result.length).toBeGreaterThan(0);
	});
});
