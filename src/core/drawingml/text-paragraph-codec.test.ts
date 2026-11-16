import { describe, expect, it } from 'vitest';

import { NS, parseXml } from '../xml/index';
import { parseDrawingParagraphProperties } from './read-paragraph-properties';
import { domTextElementView } from './text-element-view';
import {
	drawingParagraphAttributeEntries,
	drawingTabStopEntries,
	readDrawingParagraphCodecProperties,
} from './text-paragraph-codec';
import { drawingParagraphPropertiesXml } from './write-paragraph-properties';

const pPr = (xml: string) => parseXml(`<a:pPr xmlns:a="${NS.a}" ${xml}`).documentElement;

describe('DrawingML paragraph property codec', () => {
	it('reads the same families over a view as the DOM reader', () => {
		const element = pPr(
			'lvl="11" algn="ctr" marL="342900" indent="-342900" rtl="1" fontAlgn="base">' +
				'<a:lnSpc><a:spcPct val="90%"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef>' +
				'<a:buSzPct val="75000"/><a:buFont typeface=" Wingdings " charset="2"/>' +
				'<a:buChar char="§"/><a:tabLst><a:tab pos="914400" algn="l" leader="dot"/></a:tabLst>' +
				'</a:pPr>',
		);
		const codec = readDrawingParagraphCodecProperties(domTextElementView(element));
		expect(codec).toEqual({
			level: 8,
			align: 'ctr',
			marginLeftEmu: 342900,
			indentEmu: -342900,
			rightToLeft: true,
			fontAlign: 'base',
			lineSpacing: { unit: 'percent', value: 0.9 },
			spaceBefore: { unit: 'points', value: 6 },
			bulletSize: { unit: 'percent', value: 0.75 },
			bulletFont: { typeface: 'Wingdings', charset: 2 },
			bullet: { kind: 'char', char: '§' },
			tabStops: [{ positionEmu: 914400, align: 'l', leader: 'dot' }],
		});
		expect(parseDrawingParagraphProperties(element)).toEqual(codec);
	});

	it('writes attributes in schema order and keeps an authored tab leader', () => {
		expect(
			drawingParagraphAttributeEntries({
				hangingPunctuation: false,
				align: 'r',
				level: 2,
				marginLeftEmu: 457200.4,
				rightToLeft: false,
			}).map(([name]) => name),
		).toEqual(['marL', 'lvl', 'algn', 'rtl', 'hangingPunct']);
		expect(drawingTabStopEntries({ positionEmu: 12.6, leader: 'none' })).toEqual([
			['pos', '13'],
			['leader', 'none'],
		]);
		expect(
			drawingParagraphPropertiesXml({
				marginLeftEmu: 0,
				bulletSize: { unit: 'points', value: 12 },
				bullet: { kind: 'autoNumber', scheme: 'arabicPeriod', startAt: 3 },
				tabStops: [{ positionEmu: 914400, align: 'dec', leader: 'hyphen' }],
			}),
		).toBe(
			'<a:pPr marL="0" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
				'<a:buSzPts val="1200"/><a:buAutoNum type="arabicPeriod" startAt="3"/>' +
				'<a:tabLst><a:tab pos="914400" algn="dec" leader="hyphen"/></a:tabLst></a:pPr>',
		);
	});
});
