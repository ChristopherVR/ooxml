import { describe, it, expect } from 'vitest';

import { parseReferenceUri, resolveReferenceUriToPart } from './reference-utils';

describe('resolveReferenceUriToPart', () => {
	it('returns undefined for an empty string', () => {
		expect(resolveReferenceUriToPart('')).toBeUndefined();
	});

	it('returns undefined for whitespace-only input', () => {
		expect(resolveReferenceUriToPart('   ')).toBeUndefined();
	});

	it('returns undefined for a fragment-only URI', () => {
		expect(resolveReferenceUriToPart('#idPackageObject')).toBeUndefined();
	});

	it('returns undefined for a fragment-only URI with complex fragment', () => {
		expect(resolveReferenceUriToPart('#_xmlsignatures/sig1')).toBeUndefined();
	});

	it('resolves a URI with a leading slash', () => {
		expect(resolveReferenceUriToPart('/ppt/slides/slide1.xml')).toBe('ppt/slides/slide1.xml');
	});

	it('resolves a normal relative path', () => {
		expect(resolveReferenceUriToPart('ppt/slides/slide1.xml')).toBe('ppt/slides/slide1.xml');
	});

	it('keeps percent-encoding, which is part of the name', () => {
		expect(resolveReferenceUriToPart('/ppt/media/a%20b.xml')).toBe('ppt/media/a%20b.xml');
	});

	it('strips the ContentType query from a part URI', () => {
		const result = resolveReferenceUriToPart('/ppt/presentation.xml?ContentType=application/xml');
		expect(result).toBe('ppt/presentation.xml');
	});

	it.each([
		'/ppt/../secret.xml',
		'/ppt/./slide.xml',
		'/ppt//slide.xml',
		'/ppt/slide.xml#fragment',
		'/ppt/slide.xml?Other=1',
		'/ppt/slide.xml?ContentType=a?ContentType=b',
		'/ppt/%zz.xml',
		'/ppt/%41.xml',
		'/ppt%2Fslide.xml',
		'/ppt/slide 1.xml',
		'/[Content_Types].xml',
		'/ppt/slide.',
	])('rejects the ambiguous part URI %s', (uri) => {
		expect(resolveReferenceUriToPart(uri)).toBeUndefined();
	});

	it('trims surrounding whitespace before processing', () => {
		expect(resolveReferenceUriToPart('  /ppt/slides/slide1.xml  ')).toBe('ppt/slides/slide1.xml');
	});
});

describe('parseReferenceUri', () => {
	it('returns the part path and its declared content type', () => {
		expect(parseReferenceUri('/ppt/slide%201.xml?ContentType=application/xml')).toStrictEqual({
			partPath: 'ppt/slide%201.xml',
			contentType: 'application/xml',
		});
	});

	it('omits the content type when the URI has no query', () => {
		expect(parseReferenceUri('/ppt/slide1.xml')).toStrictEqual({ partPath: 'ppt/slide1.xml' });
	});
});
