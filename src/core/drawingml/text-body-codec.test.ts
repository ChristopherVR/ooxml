import { describe, expect, it } from 'vitest';

import { NS, parseXml } from '../xml/index';
import { parseDrawingBodyProperties } from './read-text-body';
import { drawingRichTextBodyXml } from './write-text-body';
import {
	drawingAutofitChild,
	drawingBodyAttributeEntries,
	readDrawingBodyCodecProperties,
} from './text-body-codec';
import { domTextElementView } from './text-element-view';

describe('DrawingML body property codec', () => {
	it('reads attributes and the autofit choice over a view', () => {
		const element = parseXml(
			`<a:bodyPr xmlns:a="${NS.a}" rot="5400000" wrap="square" lIns="91440" numCol="2" anchor="ctr" rtlCol="0">` +
				'<a:normAutofit fontScale="62.5%" lnSpcReduction="20000"/></a:bodyPr>',
		).documentElement;
		const codec = readDrawingBodyCodecProperties(domTextElementView(element));
		expect(codec).toEqual({
			rotation: 90,
			wrap: 'square',
			insetLeftEmu: 91440,
			columnCount: 2,
			anchor: 'ctr',
			rightToLeftColumns: false,
			autofit: { kind: 'normal', fontScale: 0.625, lineSpacingReduction: 0.2 },
		});
		expect(parseDrawingBodyProperties(element)).toEqual(codec);
	});

	it('reads the text warp and keeps a lossy one verbatim', () => {
		const body = (warp: string) =>
			parseDrawingBodyProperties(
				parseXml(`<a:bodyPr xmlns:a="${NS.a}">${warp}</a:bodyPr>`).documentElement,
			);
		expect(
			body(
				'<a:prstTxWarp prst="textWave1"><a:avLst><a:gd name="adj1" fmla="val 12500"/></a:avLst></a:prstTxWarp>',
			),
		).toEqual({
			textWarp: { preset: 'textWave1', guides: [{ name: 'adj1', formula: 'val 12500' }] },
		});
		expect(body('<a:prstTxWarp prst="textPlain"/>')).toEqual({ textWarp: { preset: 'textPlain' } });
		expect(
			body('<a:prstTxWarp prst="textPlain"><a:extLst/></a:prstTxWarp>')?.extraXml,
		).toHaveLength(1);
		const written = drawingRichTextBodyXml({
			bodyProperties: { textWarp: { preset: 'textArchUp', guides: [] }, wrap: 'none' },
			paragraphs: [],
		});
		expect(written).toContain('<a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp>');
	});

	it('writes schema-ordered attributes and the autofit child', () => {
		expect(
			drawingBodyAttributeEntries({ anchor: 'b', insetTopEmu: 45720.4, rotation: -45 }),
		).toEqual([
			['rot', '-2700000'],
			['tIns', '45720'],
			['anchor', 'b'],
		]);
		expect(drawingAutofitChild({ kind: 'normal', fontScale: 0.9 })).toEqual({
			local: 'normAutofit',
			attributes: [['fontScale', '90000']],
		});
		expect(drawingAutofitChild({ kind: 'shape' })).toEqual({ local: 'spAutoFit', attributes: [] });
	});
});
