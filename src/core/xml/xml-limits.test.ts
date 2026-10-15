import { describe, expect, it } from 'vitest';
import { NS, parseXml, textContent } from './index';

describe('parseXml resource limits', () => {
	it('rejects nesting deeper than the default depth limit', () => {
		const depth = 10_000;
		const xml = `<root>${'<a>'.repeat(depth)}${'</a>'.repeat(depth)}</root>`;
		expect(() => parseXml(xml, { label: 'DOCX' })).toThrow(/DOCX XML exceeds the nesting depth/);
	});

	it('accepts nesting up to a custom limit and rejects beyond it', () => {
		const xml = `<root>${'<a>'.repeat(100)}${'</a>'.repeat(100)}</root>`;
		expect(parseXml(xml).documentElement.localName).toBe('root');
		expect(() => parseXml(xml, { limits: { maxDepth: 50 } })).toThrow(/depth limit of 50/);
	});

	it('rejects a part with more elements than the node limit', () => {
		const xml = `<root>${'<a/>'.repeat(20)}</root>`;
		expect(() => parseXml(xml, { limits: { maxNodes: 10 } })).toThrow(/limit of 10 elements/);
		expect(parseXml(xml, { limits: { maxNodes: 21 } }).documentElement.childNodes.length).toBe(20);
	});

	it('does not count markup inside comments, CDATA or attribute values', () => {
		const inner = '<a>'.repeat(600);
		const xml = `<root><!-- ${inner} --><t><![CDATA[${inner}]]></t><e v="&gt;&gt;" w='x>y'/></root>`;
		expect(parseXml(xml, { limits: { maxDepth: 4 } }).documentElement.localName).toBe('root');
	});

	it('still parses ordinary documents within the defaults', () => {
		const xml = `<w:document xmlns:w="${NS.w}"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`;
		expect(textContent(parseXml(xml).documentElement)).toBe('hi');
	});
});
