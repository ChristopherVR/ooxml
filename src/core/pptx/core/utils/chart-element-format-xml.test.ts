import { describe, expect, it } from 'vitest';

import type { XmlObject } from '../types';
import { applyChartElementFormatToSpPr } from './chart-element-format-xml';

const local = (key: string) => key.replace(/^.*:/u, '');

describe('applyChartElementFormatToSpPr', () => {
	it('adds transparency to an authored theme colour without replacing it', () => {
		const spPr: XmlObject = { 'a:solidFill': { 'a:schemeClr': { '@_val': 'accent2' } } };
		const out = applyChartElementFormatToSpPr(
			spPr,
			{ fill: { kind: 'solid', transparency: 40 } },
			local,
		);
		expect(out).toEqual({
			'a:solidFill': { 'a:schemeClr': { '@_val': 'accent2', 'a:alpha': { '@_val': '60000' } } },
		});
		expect(spPr['a:solidFill']).toEqual({ 'a:schemeClr': { '@_val': 'accent2' } });
	});

	it('puts the fill before the line and keeps passthrough children in order', () => {
		const spPr: XmlObject = { 'a:ln': { '@_w': '9525' }, 'a:scene3d': {}, 'a:extLst': {} };
		const out = applyChartElementFormatToSpPr(spPr, { fill: { kind: 'none' } }, local)!;
		expect(Object.keys(out)).toEqual(['a:noFill', 'a:ln', 'a:scene3d', 'a:extLst']);
	});

	it('removes the fill for Automatic and the whole node when nothing is left', () => {
		const spPr: XmlObject = { 'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } } };
		expect(
			applyChartElementFormatToSpPr(spPr, { fill: { kind: 'automatic' } }, local),
		).toBeUndefined();
	});

	it('replaces the join and writes arrows in CT_LineProperties order', () => {
		const spPr: XmlObject = {
			'a:ln': {
				'a:miter': { '@_lim': '800000' },
				'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
			},
		};
		const out = applyChartElementFormatToSpPr(
			spPr,
			{ line: { kind: 'solid', join: 'bevel', beginArrow: 'oval', dash: 'sysDot' } },
			local,
		)!;
		const ln = out['a:ln'] as XmlObject;
		expect(Object.keys(ln).filter((k) => !k.startsWith('@_'))).toEqual([
			'a:solidFill',
			'a:prstDash',
			'a:bevel',
			'a:headEnd',
		]);
	});

	it('removes a single effect with null and drops an emptied effect list', () => {
		const spPr: XmlObject = {
			'a:effectLst': { 'a:glow': { '@_rad': '1' }, 'a:softEdge': { '@_rad': '2' } },
		};
		const out = applyChartElementFormatToSpPr(spPr, { effects: { glow: null } }, local)!;
		expect(out['a:effectLst']).toEqual({ 'a:softEdge': { '@_rad': '2' } });
		expect(
			applyChartElementFormatToSpPr(out, { effects: { softEdge: null } }, local),
		).toBeUndefined();
	});

	it('Reset drops the authored node before applying later edits', () => {
		const spPr: XmlObject = { 'a:solidFill': {}, 'a:effectLst': { 'a:glow': {} } };
		expect(applyChartElementFormatToSpPr(spPr, { reset: true }, local)).toBeUndefined();
		expect(
			applyChartElementFormatToSpPr(spPr, { reset: true, line: { kind: 'none' } }, local),
		).toEqual({ 'a:ln': { 'a:noFill': {} } });
	});
});
