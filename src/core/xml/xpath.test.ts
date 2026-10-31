import { describe, expect, it } from 'vitest';
import { parseXml, selectXmlNodes } from './index';

describe('XPath node selection', () => {
	it('selects namespaced nodes with predicates and axes', () => {
		const doc = parseXml('<r xmlns:p="urn:p"><p:a id="1"/><p:a id="2"/></r>');
		expect(
			selectXmlNodes('//p:a[@id="2"]/preceding-sibling::p:a/@id', doc, {
				namespaces: { p: 'urn:p' },
			}).map((node) => node.nodeValue),
		).toEqual(['1']);
	});
	it('makes default-namespace matching explicit', () => {
		const doc = parseXml('<r xmlns="urn:p"><a>yes</a></r>');
		expect(selectXmlNodes('//a', doc)).toEqual([]);
		expect(selectXmlNodes('//a', doc, { matchDefaultNamespace: true })[0]?.textContent).toBe('yes');
	});
	it('rejects scalar queries and recoverable XML syntax in strict mode', () => {
		expect(() => selectXmlNodes('count(//a)', parseXml('<r/>'))).toThrow();
		expect(() => parseXml('<r a=unquoted/>', { warningsAsErrors: true })).toThrow();
		expect(parseXml('<r a=unquoted/>').documentElement.getAttribute('a')).toBe('unquoted');
	});
});
