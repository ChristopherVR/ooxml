import { DOMParser } from '@xmldom/xmldom';
import { describe, it, expect } from 'vitest';

import { localNameOf, firstDescendantNamed } from './dom';

/** Parses with xmldom, typed as the structural document the module accepts (as its callers do). */
const parseXml = (xml: string): Document =>
	new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document;

// ---------------------------------------------------------------------------
// localNameOf
// ---------------------------------------------------------------------------

describe('localNameOf', () => {
	it('returns local name stripped of namespace prefix', () => {
		const doc = parseXml('<ds:Signature xmlns:ds="http://www.w3.org/2000/09/xmldsig#"/>');
		expect(localNameOf(doc.documentElement!)).toBe('Signature');
	});

	it('returns the full name when no prefix is present', () => {
		const doc = parseXml('<Root/>');
		expect(localNameOf(doc.documentElement!)).toBe('Root');
	});

	it('handles deeply-prefixed names like a:b:Foo by stripping first prefix', () => {
		const doc = parseXml('<ns:Element xmlns:ns="urn:test"/>');
		expect(localNameOf(doc.documentElement!)).toBe('Element');
	});
});

// ---------------------------------------------------------------------------
// firstDescendantNamed
// ---------------------------------------------------------------------------

describe('firstDescendantNamed', () => {
	it('finds element by local name in parsed XML', () => {
		const doc = parseXml(
			'<Root><ds:SignedInfo xmlns:ds="http://www.w3.org/2000/09/xmldsig#"><ds:Reference/></ds:SignedInfo></Root>',
		);
		const result = firstDescendantNamed(doc, 'SignedInfo');
		expect(result).toBeDefined();
		expect(localNameOf(result!)).toBe('SignedInfo');
	});

	it('returns the first matching element when multiple exist', () => {
		const doc = parseXml('<Root><Item id="1"/><Item id="2"/></Root>');
		const result = firstDescendantNamed(doc, 'Item');
		expect(result).toBeDefined();
		expect(result!.getAttribute('id')).toBe('1');
	});

	it('returns undefined when element is not found', () => {
		const doc = parseXml('<Root><Child/></Root>');
		const result = firstDescendantNamed(doc, 'NonExistent');
		expect(result).toBeUndefined();
	});

	it('searches within an element subtree', () => {
		const doc = parseXml('<Root><Parent><Target/></Parent></Root>');
		const parent = firstDescendantNamed(doc, 'Parent')!;
		const result = firstDescendantNamed(parent, 'Target');
		expect(result).toBeDefined();
		expect(localNameOf(result!)).toBe('Target');
	});
});
