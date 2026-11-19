import { describe, expect, it } from 'vitest';

import { isPartName, partNameKey } from './part-names';

describe('partNameKey', () => {
	it('treats names as equivalent under ASCII case-insensitive matching', () => {
		expect(partNameKey('/PPT/Slides/Slide1.XML')).toBe(partNameKey('ppt/slides/slide1.xml'));
	});

	it('folds only ASCII letters', () => {
		expect(partNameKey('/É.xml')).not.toBe(partNameKey('/é.xml'));
		expect(partNameKey('/K.xml')).not.toBe(partNameKey('/k.xml'));
	});

	it('maps a percent-encoded non-ASCII ZIP item name to its part name', () => {
		expect(partNameKey('media/%C3%A9.png')).toBe(partNameKey('/media/é.png'));
	});

	it('keeps ASCII percent-encoding, which is part of the name', () => {
		expect(partNameKey('a%20b.xml')).not.toBe(partNameKey('a b.xml'));
		expect(partNameKey('a%2Fb.xml')).toBe(partNameKey('a%2fb.xml'));
	});

	it('leaves octets that are not UTF-8 encoded', () => {
		expect(partNameKey('a%FF.xml')).toBe('a%ff.xml');
	});
});

describe('isPartName', () => {
	it.each([
		'/ppt/slides/slide1.xml',
		'ppt/slides/slide1.xml',
		'/_rels/.rels',
		'/ppt/media/a%20b.xml',
		'/media/%C3%A9.png',
		'/media/é.png',
	])('accepts %s', (name) => {
		expect(isPartName(name)).toBe(true);
	});

	it.each([
		['an empty name', ''],
		['an empty segment', '/ppt//slide.xml'],
		['a trailing slash', '/ppt/'],
		['a segment ending in a dot', '/ppt/slide.'],
		['a dot segment', '/ppt/../slide.xml'],
		['an encoded unreserved character', '/ppt/slide%31.xml'],
		['an encoded slash', '/ppt%2Fslide.xml'],
		['an encoded backslash', '/ppt%5Cslide.xml'],
		['a raw space', '/ppt/a b.xml'],
		['brackets', '/[Content_Types].xml'],
		['a malformed escape', '/ppt/%zz.xml'],
	])('rejects %s', (_label, name) => {
		expect(isPartName(name)).toBe(false);
	});
});
