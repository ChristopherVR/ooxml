import { describe, expect, it } from 'vitest';

import { PptxXmlLookupService } from '../services/PptxXmlLookupService';
import type { XmlObject } from '../types';
import {
	buildDefRPrTextProperties,
	parseDefRPrTextStyle,
	parseRichTextStyle,
} from './chart-def-rpr-style';

const xmlLookup = new PptxXmlLookupService();
const colorParser = { parseColor: () => undefined };

describe('chart a:defRPr East Asian typeface', () => {
	it('reads a:ea next to a:latin', () => {
		const defRPr: XmlObject = {
			'a:latin': { '@_typeface': 'Arial' },
			'a:ea': { '@_typeface': '+mn-ea' },
		};
		const resolve = (raw: string) => (raw === '+mn-ea' ? 'Malgun Gothic' : raw);
		expect(parseDefRPrTextStyle(defRPr, xmlLookup, colorParser, resolve)).toStrictEqual({
			fontFamily: 'Arial',
			eastAsiaFontFamily: 'Malgun Gothic',
		});
	});

	it('reads a label that names only an East Asian face', () => {
		const defRPr: XmlObject = { 'a:ea': { '@_typeface': 'Malgun Gothic' } };
		expect(parseDefRPrTextStyle(defRPr, xmlLookup, colorParser)).toStrictEqual({
			eastAsiaFontFamily: 'Malgun Gothic',
		});
	});

	it('writes back an authored a:ea unchanged', () => {
		const authored: XmlObject = { 'a:ea': { '@_typeface': '+mn-ea' } };
		const txPr = buildDefRPrTextProperties(
			{ fontSize: 9, eastAsiaFontFamily: 'Malgun Gothic' },
			authored,
		);
		const paragraph = txPr?.['a:p'] as XmlObject;
		const rPr = (paragraph['a:pPr'] as XmlObject)['a:defRPr'] as XmlObject;
		expect(rPr['a:ea']).toStrictEqual({ '@_typeface': '+mn-ea' });
	});

	it('writes a:ea after a:latin', () => {
		const txPr = buildDefRPrTextProperties(
			{ fontFamily: 'Arial', eastAsiaFontFamily: 'Malgun Gothic' },
			undefined,
		);
		const paragraph = txPr?.['a:p'] as XmlObject;
		const rPr = (paragraph['a:pPr'] as XmlObject)['a:defRPr'] as XmlObject;
		expect(Object.keys(rPr)).toStrictEqual(['a:latin', 'a:ea']);
		expect(rPr['a:ea']).toStrictEqual({ '@_typeface': 'Malgun Gothic' });
	});
});

describe('parseRichTextStyle', () => {
	it("reads the first run's a:rPr over the paragraph's a:defRPr", () => {
		const rich: XmlObject = {
			'a:p': {
				'a:pPr': { 'a:defRPr': { '@_sz': '900', '@_b': '1' } },
				'a:r': { 'a:rPr': { '@_sz': '800' }, 'a:t': '1,444' },
			},
		};
		expect(parseRichTextStyle(rich, xmlLookup, colorParser)).toStrictEqual({
			fontSize: 8,
			bold: true,
		});
	});

	it('reads a label written only as a field', () => {
		const rich: XmlObject = {
			'a:p': { 'a:fld': { '@_type': 'VALUE', 'a:rPr': { '@_sz': '800' }, 'a:t': '[VALUE]' } },
		};
		expect(parseRichTextStyle(rich, xmlLookup, colorParser)).toStrictEqual({ fontSize: 8 });
	});
});
