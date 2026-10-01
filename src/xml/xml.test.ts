import { describe, expect, it } from 'vitest';
import {
	NS,
	attr,
	buildXml,
	children,
	elements,
	first,
	makeNS,
	named,
	parseXml,
	relAttr,
	textContent,
} from './index.js';

const DOC = `<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:body><w:p w:rsidR="1"><w:r><w:t>Hi</w:t></w:r></w:p><w:p/><x:other xmlns:x="urn:x"/></w:body></w:document>`;

describe('parseXml', () => {
	it('parses a part and exposes the DOM', () => {
		const doc = parseXml(DOC);
		expect(doc.documentElement.localName).toBe('document');
		expect(buildXml(doc)).toBe(DOC);
	});

	it('rejects DTD and entity declarations before parsing', () => {
		expect(() => parseXml('<!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>')).toThrow(
			'OOXML XML with DTD or entity declarations is not supported',
		);
		expect(() => parseXml('<!ENTITY x "y"><a/>', { label: 'DOCX' })).toThrow(
			'DOCX XML with DTD or entity declarations is not supported',
		);
	});

	it('throws on malformed input and on a missing root, naming the document kind', () => {
		expect(() => parseXml('<a><b></a>')).toThrow(/^Invalid OOXML XML:/);
		expect(() => parseXml('', { label: 'PPTX' })).toThrow(/^Invalid PPTX XML:/);
		expect(() => parseXml('not xml at all')).toThrow(/^Invalid OOXML XML:/);
	});
});

describe('namespace-aware helpers', () => {
	const body = first(parseXml(DOC).documentElement, 'body', NS.w)!;

	it('finds children by local name in a namespace and ignores other namespaces', () => {
		expect(children(body, 'p', NS.w)).toHaveLength(2);
		expect(children(body, 'other', NS.w)).toHaveLength(0);
		expect(children(body, 'other', 'urn:x')).toHaveLength(1);
		expect(elements(body)).toHaveLength(3);
		expect(first(body, 'missing', NS.w)).toBeUndefined();
		expect(first(undefined, 'p', NS.w)).toBeUndefined();
	});

	it('matches un-namespaced elements only when lenient', () => {
		const bare = parseXml('<root><p/></root>').documentElement;
		expect(children(bare, 'p', NS.w)).toHaveLength(1);
		expect(children(bare, 'p', NS.w, false)).toHaveLength(0);
		expect(named(bare, 'root', NS.w)).toBe(true);
		expect(named(null, 'root', NS.w)).toBe(false);
	});

	it('reads attributes by namespace with a prefix fallback', () => {
		const para = children(body, 'p', NS.w)[0]!;
		expect(attr(para, 'rsidR', NS.w, 'w')).toBe('1');
		expect(attr(para, 'absent', NS.w, 'w')).toBeFalsy();
		// Attributes added with setAttribute('w:val', ...) have no namespace; the prefix form finds them.
		const bare = parseXml('<e/>').documentElement;
		bare.setAttribute('w:val', '2');
		expect(attr(bare, 'val', NS.w, 'w')).toBe('2');
		const embed = parseXml(`<e xmlns:r="${NS.r}" r:embed="rId9"/>`).documentElement;
		expect(relAttr(embed, 'embed')).toBe('rId9');
	});

	it('creates namespaced elements that serialize with their prefix', () => {
		const doc = parseXml('<root/>');
		const el = makeNS(doc, NS.a, 'a:ln');
		doc.documentElement.appendChild(el);
		expect(textContent(el)).toBe('');
		expect(buildXml(doc)).toContain(`<a:ln xmlns:a="${NS.a}"/>`);
	});
});
