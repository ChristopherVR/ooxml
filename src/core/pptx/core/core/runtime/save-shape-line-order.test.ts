import { describe, expect, it } from 'vitest';

import type { ShapeStyle, XmlObject } from '../../types';
import { writeShapeStroke } from './save-shape-fill-stroke';

/**
 * `writeShapeStroke` adds `a:ln` children as each writer runs (arrow ends,
 * then the join), and fast-xml-parser emits keys in insertion order. The
 * written node must still follow CT_LineProperties (ECMA-376 §20.1.2.2.24):
 * fill, dash, join, `headEnd`, `tailEnd`.
 */

const ctx = {
	emuPerPx: 9525,
	parseColor: (): string | undefined => undefined,
};

const childKeys = (node: XmlObject): string[] =>
	Object.keys(node).filter((key) => !key.startsWith('@_'));

describe('writeShapeStroke child order', () => {
	it('writes the join before the ends and the head before the tail', () => {
		const spPr: XmlObject = {};
		const style: ShapeStyle = {
			strokeColor: '#FF0000',
			strokeWidth: 2,
			strokeDash: 'dash',
			lineJoin: 'round',
			connectorStartArrow: 'triangle',
			connectorEndArrow: 'oval',
		};
		writeShapeStroke(spPr, style, ctx);
		expect(childKeys(spPr['a:ln'] as XmlObject)).toEqual([
			'a:solidFill',
			'a:prstDash',
			'a:round',
			'a:headEnd',
			'a:tailEnd',
		]);
	});

	it('reorders a connector line whose ends already exist', () => {
		const spPr: XmlObject = {
			'a:ln': {
				'@_w': '9525',
				'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
				'a:headEnd': { '@_type': 'none' },
				'a:tailEnd': { '@_type': 'none' },
			},
		};
		writeShapeStroke(
			spPr,
			{ lineJoin: 'miter', miterLimit: 800000, connectorEndArrow: 'arrow' },
			ctx,
		);
		const line = spPr['a:ln'] as XmlObject;
		expect(childKeys(line)).toEqual(['a:solidFill', 'a:miter', 'a:headEnd', 'a:tailEnd']);
		expect(line['a:tailEnd']).toEqual({ '@_type': 'arrow' });
	});

	it('leaves a line already in schema order as it was', () => {
		const line: XmlObject = {
			'@_w': '12700',
			'a:solidFill': { 'a:srgbClr': { '@_val': '00FF00' } },
			'a:custDash': { 'a:ds': { '@_d': '100000', '@_sp': '50000' } },
			'a:bevel': {},
			'a:extLst': {},
		};
		const spPr: XmlObject = { 'a:ln': { ...line } };
		writeShapeStroke(spPr, {}, ctx);
		expect(spPr['a:ln']).toEqual(line);
		expect(childKeys(spPr['a:ln'] as XmlObject)).toEqual(childKeys(line));
	});
});
