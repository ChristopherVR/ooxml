import { describe, expect, it } from 'vitest';

import { NS, parseXml } from '../xml/index';
import { parseDrawingRunProperties } from './read-run-properties';
import { domTextElementView } from './text-element-view';
import { drawingRunAttributeEntries, readDrawingRunCodecProperties } from './text-run-codec';

describe('DrawingML run property codec', () => {
	it('reads attributes, fonts, rtl and the follow-text markers over a view', () => {
		const element = parseXml(
			`<a:rPr xmlns:a="${NS.a}" lang="en-US" sz="1800" b="true" u="dbl" baseline="30%" spc="-50" dirty="0">` +
				'<a:uLnTx/><a:latin typeface=" +mn-lt " pitchFamily="34"/><a:sym typeface="Wingdings"/>' +
				'<a:rtl/></a:rPr>',
		).documentElement;
		const codec = readDrawingRunCodecProperties(domTextElementView(element));
		expect(codec).toEqual({
			language: 'en-US',
			sizePt: 18,
			bold: true,
			underline: 'dbl',
			baseline: 0.3,
			spacingPt: -0.5,
			dirty: false,
			underlineLineFollowsText: true,
			latin: { typeface: '+mn-lt', pitchFamily: 34 },
			symbol: { typeface: 'Wingdings' },
			rightToLeft: true,
		});
		expect(parseDrawingRunProperties(element)).toEqual(codec);
	});

	it('writes the attributes in schema order', () => {
		expect(
			drawingRunAttributeEntries({
				dirty: false,
				bold: true,
				sizePt: 10.5,
				language: 'de-DE',
				kumimoji: true,
				baseline: -0.25,
			}),
		).toEqual([
			['kumimoji', '1'],
			['lang', 'de-DE'],
			['sz', '1050'],
			['b', '1'],
			['baseline', '-25000'],
			['dirty', '0'],
		]);
	});
});
