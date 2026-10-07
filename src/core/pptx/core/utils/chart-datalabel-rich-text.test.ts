import { describe, expect, it } from 'vitest';

import type { XmlObject } from '../types';
import { buildDataLabelTx } from './chart-datalabel-rich-text';

const getLocalName = (key: string): string => key.slice(key.indexOf(':') + 1);

const rPr: XmlObject = { '@_sz': '800', '@_b': '1', 'a:latin': { '@_typeface': 'Georgia' } };

function tx(paragraph: XmlObject): XmlObject {
	return { 'c:rich': { 'a:bodyPr': {}, 'a:lstStyle': {}, 'a:p': paragraph } };
}

const paragraphOf = (node: XmlObject): XmlObject =>
	(node['c:rich'] as XmlObject)['a:p'] as XmlObject;

describe('buildDataLabelTx', () => {
	it('returns the authored c:tx when the text is unchanged', () => {
		const authored = tx({
			'a:r': { 'a:rPr': rPr, 'a:t': { '@_xml:space': 'preserve', '#text': '1,444' } },
		});
		expect(buildDataLabelTx(authored, '1,444', getLocalName)).toBe(authored);
	});

	it("keeps the first run's formatting, a:pPr and a:endParaRPr when the text changes", () => {
		const endParaRPr = { '@_lang': 'en-US' };
		const authored = tx({
			'a:pPr': { 'a:defRPr': {} },
			'a:r': [
				{ 'a:rPr': rPr, 'a:t': '1,' },
				{ 'a:rPr': { '@_sz': '600' }, 'a:t': '444' },
			],
			'a:endParaRPr': endParaRPr,
		});
		expect(paragraphOf(buildDataLabelTx(authored, '1,500', getLocalName))).toStrictEqual({
			'a:pPr': { 'a:defRPr': {} },
			'a:r': { 'a:rPr': rPr, 'a:t': '1,500' },
			'a:endParaRPr': endParaRPr,
		});
	});

	it('takes the formatting of a field when the label has no run', () => {
		const authored = tx({ 'a:fld': { '@_type': 'VALUE', 'a:rPr': rPr, 'a:t': '[VALUE]' } });
		expect(paragraphOf(buildDataLabelTx(authored, '1,500', getLocalName))).toStrictEqual({
			'a:r': { 'a:rPr': rPr, 'a:t': '1,500' },
		});
	});

	it('writes a plain c:rich for a label that had none', () => {
		expect(buildDataLabelTx(undefined, 'Peak', getLocalName)).toStrictEqual({
			'c:rich': { 'a:bodyPr': {}, 'a:lstStyle': {}, 'a:p': { 'a:r': { 'a:t': 'Peak' } } },
		});
	});
});
