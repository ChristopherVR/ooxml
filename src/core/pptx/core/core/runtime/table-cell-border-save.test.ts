/**
 * The edge lines of a cell written over the ones it already has: a preset
 * and a custom dash replace each other in place, in `CT_LineProperties`
 * order, and a custom dash the model has no pattern for is left alone.
 */
import { describe, expect, it } from 'vitest';

import type { XmlObject } from '../../types';
import { writeCellEdgeBorders } from './table-cell-border-save';

const EMU_PER_PX = 9525;

const elementKeys = (node: unknown): string[] =>
	Object.keys(node as XmlObject).filter((key) => !key.startsWith('@_'));

describe('writeCellEdgeBorders', () => {
	it('replaces a preset dash with a custom one, before the join and line ends', () => {
		const tcPr: XmlObject = {
			'a:lnL': {
				'@_w': '12700',
				'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
				'a:prstDash': { '@_val': 'dash' },
				'a:round': {},
				'a:headEnd': { '@_type': 'none' },
			},
		};
		writeCellEdgeBorders(
			tcPr,
			{ borderLeftDash: 'custom', borderLeftDashSegments: [{ dash: 300000, space: 100000 }] },
			EMU_PER_PX,
		);
		const ln = tcPr['a:lnL'] as XmlObject;
		expect(elementKeys(ln)).toEqual(['a:solidFill', 'a:custDash', 'a:round', 'a:headEnd']);
		expect(ln['a:custDash']).toEqual({ 'a:ds': { '@_d': '300000', '@_sp': '100000' } });
		expect(ln['@_w']).toBe('12700');
	});

	it('replaces a custom dash with a preset, and removes either for solid', () => {
		const custom = (): XmlObject => ({
			'a:lnT': { 'a:custDash': { 'a:ds': { '@_d': '100000', '@_sp': '100000' } } },
		});
		const preset = custom();
		writeCellEdgeBorders(preset, { borderTopDash: 'sysDash' }, EMU_PER_PX);
		expect(preset['a:lnT']).toEqual({ 'a:prstDash': { '@_val': 'sysDash' } });

		const solid = custom();
		writeCellEdgeBorders(solid, { borderTopDash: 'solid' }, EMU_PER_PX);
		expect(solid['a:lnT']).toEqual({});
	});

	it('leaves a custom dash with no pattern in the model as it is', () => {
		const ds = { '@_d': '-5', '@_sp': 'x' };
		const tcPr: XmlObject = { 'a:lnB': { 'a:custDash': { 'a:ds': ds } } };
		writeCellEdgeBorders(tcPr, { borderBottomDash: 'custom' }, EMU_PER_PX);
		expect(tcPr['a:lnB']).toEqual({ 'a:custDash': { 'a:ds': ds } });
	});

	it('puts a colour added to a dashed line before its dash, replacing a:noFill', () => {
		const tcPr: XmlObject = {
			'a:lnR': { 'a:noFill': {}, 'a:prstDash': { '@_val': 'dot' } },
		};
		writeCellEdgeBorders(tcPr, { borderRightColor: '#FF0000' }, EMU_PER_PX);
		expect(elementKeys(tcPr['a:lnR'])).toEqual(['a:solidFill', 'a:prstDash']);
	});
});
