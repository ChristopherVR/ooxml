import { describe, expect, it } from 'vitest';

import type { PptxChartGradientFill, XmlObject } from '../types';
import {
	applyChartGradientToSpPr,
	buildChartGradFillXml,
	chartGradientsEqual,
	supportsChartGradientFill,
} from './chart-gradient-fill-writer';

const localName = (key: string) => key.replace(/^.*:/u, '');

const linear: PptxChartGradientFill = {
	type: 'linear',
	angle: 90,
	stops: [
		{ color: '#60A5FA', position: 0 },
		{ color: '#1E3A8A', position: 100 },
	],
};

/** Stand-in for the load-time parser: reads back what these tests author. */
function fakeParse(gradient: PptxChartGradientFill) {
	return (spPr: XmlObject | undefined) => (spPr?.['a:gradFill'] ? gradient : undefined);
}

describe('buildChartGradFillXml', () => {
	it('writes a fresh linear gradient in OOXML units', () => {
		expect(buildChartGradFillXml(linear)).toStrictEqual({
			'@_rotWithShape': '1',
			'a:gsLst': {
				'a:gs': [
					{ '@_pos': '0', 'a:srgbClr': { '@_val': '60A5FA' } },
					{ '@_pos': '100000', 'a:srgbClr': { '@_val': '1E3A8A' } },
				],
			},
			'a:lin': { '@_ang': '5400000', '@_scaled': '0' },
		});
	});

	it('sorts stops, normalises the angle and writes stop opacity as a:alpha', () => {
		const xml = buildChartGradFillXml({
			type: 'linear',
			angle: -45,
			stops: [
				{ color: '#000000', position: 100 },
				{ color: 'ffffff', position: 25.5, opacity: 0.4 },
			],
		});
		expect(xml['a:lin']).toStrictEqual({ '@_ang': String(315 * 60000), '@_scaled': '0' });
		expect((xml['a:gsLst'] as XmlObject)['a:gs']).toStrictEqual([
			{ '@_pos': '25500', 'a:srgbClr': { '@_val': 'FFFFFF', 'a:alpha': { '@_val': '40000' } } },
			{ '@_pos': '100000', 'a:srgbClr': { '@_val': '000000' } },
		]);
	});

	it('writes a radial gradient as a:path circle with a fillToRect around the focal point', () => {
		const xml = buildChartGradFillXml({
			type: 'radial',
			focalPoint: { x: 0.25, y: 0.75 },
			stops: linear.stops,
		});
		expect(xml['a:lin']).toBeUndefined();
		expect(xml['a:path']).toStrictEqual({
			'@_path': 'circle',
			'a:fillToRect': { '@_l': '25000', '@_t': '75000', '@_r': '75000', '@_b': '25000' },
		});
	});

	it('centres a radial gradient without a focal point', () => {
		const xml = buildChartGradFillXml({ type: 'radial', stops: linear.stops });
		expect((xml['a:path'] as XmlObject)['a:fillToRect']).toStrictEqual({
			'@_l': '50000',
			'@_t': '50000',
			'@_r': '50000',
			'@_b': '50000',
		});
	});
});

describe('chartGradientsEqual', () => {
	it('compares at OOXML resolution with parser defaults', () => {
		expect(
			chartGradientsEqual(linear, {
				type: 'linear',
				angle: 450,
				stops: [
					{ color: '#1e3a8a', position: 100, opacity: 1 },
					{ color: '60a5fa', position: 0.0001 },
				],
			}),
		).toBeTruthy();
		expect(chartGradientsEqual({ ...linear, angle: undefined }, linear)).toBeTruthy();
		expect(chartGradientsEqual({ ...linear, angle: 0 }, linear)).toBeFalsy();
		expect(
			chartGradientsEqual(
				{ type: 'radial', stops: linear.stops },
				{ type: 'radial', stops: linear.stops, focalPoint: { x: 0.5, y: 0.5 } },
			),
		).toBeTruthy();
		expect(chartGradientsEqual({ type: 'radial', stops: linear.stops }, linear)).toBeFalsy();
	});
});

describe('applyChartGradientToSpPr', () => {
	it('replaces the fill choice and keeps CT_ShapeProperties order (fill before a:ln)', () => {
		const ln = { '@_w': '9525', 'a:solidFill': { 'a:schemeClr': { '@_val': 'lt1' } } };
		const spPr: XmlObject = {
			'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } },
			'a:ln': ln,
			'a:effectLst': {},
			'a:scene3d': {},
			'a:sp3d': {},
		};
		expect(applyChartGradientToSpPr(spPr, linear, localName)).toBeTruthy();
		expect(Object.keys(spPr)).toStrictEqual([
			'a:gradFill',
			'a:ln',
			'a:effectLst',
			'a:scene3d',
			'a:sp3d',
		]);
		expect(spPr['a:ln']).toBe(ln);
	});

	it('inserts after a:xfrm/a:prstGeom and drops a:noFill', () => {
		const spPr: XmlObject = { 'a:xfrm': {}, 'a:noFill': {}, 'a:ln': {} };
		applyChartGradientToSpPr(spPr, linear, localName);
		expect(Object.keys(spPr)).toStrictEqual(['a:xfrm', 'a:gradFill', 'a:ln']);
	});

	it('leaves an authored gradient the model still describes untouched', () => {
		const authored = {
			'@_flip': 'none',
			'a:gsLst': { 'a:gs': [] },
			'a:lin': { '@_ang': '5400000' },
		};
		const spPr: XmlObject = { 'a:gradFill': authored, 'a:ln': {} };
		const changed = applyChartGradientToSpPr(spPr, linear, localName, {
			parseGradient: fakeParse(linear),
		});
		expect(changed).toBeFalsy();
		expect(spPr['a:gradFill']).toBe(authored);
	});

	it('removes a parseable authored gradient the model dropped', () => {
		const spPr: XmlObject = { 'a:gradFill': {}, 'a:ln': {} };
		expect(
			applyChartGradientToSpPr(spPr, undefined, localName, { parseGradient: fakeParse(linear) }),
		).toBeTruthy();
		expect(Object.keys(spPr)).toStrictEqual(['a:ln']);
	});

	it('never removes a gradient it cannot parse', () => {
		const spPr: XmlObject = { 'a:gradFill': {} };
		expect(
			applyChartGradientToSpPr(spPr, undefined, localName, { parseGradient: () => undefined }),
		).toBeFalsy();
		expect(spPr['a:gradFill']).toBeDefined();
	});

	it('ignores line-drawn series entirely', () => {
		const spPr: XmlObject = { 'a:ln': { 'a:solidFill': {} } };
		expect(applyChartGradientToSpPr(spPr, linear, localName, { lineDrawn: true })).toBeFalsy();
		expect(spPr).toStrictEqual({ 'a:ln': { 'a:solidFill': {} } });
	});

	it('keeps an unchanged themed stop and authored attributes when another stop is edited', () => {
		const themedStop = { 'a:schemeClr': { '@_val': 'accent1', 'a:lumMod': { '@_val': '60000' } } };
		const authored: XmlObject = {
			'@_rotWithShape': '0',
			'a:gsLst': {
				'a:gs': [
					{ '@_pos': '0', ...themedStop },
					{ '@_pos': '100000', 'a:schemeClr': { '@_val': 'accent2' } },
				],
			},
			'a:lin': { '@_ang': '5400000', '@_scaled': '1' },
			'a:tileRect': {},
		};
		const parsed: PptxChartGradientFill = {
			type: 'linear',
			angle: 90,
			stops: [
				{ color: '#8FAADC', position: 0 },
				{ color: '#ED7D31', position: 100 },
			],
		};
		const spPr: XmlObject = { 'a:gradFill': authored };
		applyChartGradientToSpPr(
			spPr,
			{ ...parsed, angle: 0, stops: [parsed.stops[0], { color: '#000000', position: 100 }] },
			localName,
			{
				parseGradient: () => parsed,
				resolveColor: (node) => (node['a:schemeClr'] ? '#8FAADC' : undefined),
			},
		);
		expect(spPr['a:gradFill']).toStrictEqual({
			'@_rotWithShape': '0',
			'a:gsLst': {
				'a:gs': [
					{ '@_pos': '0', ...themedStop },
					{ '@_pos': '100000', 'a:srgbClr': { '@_val': '000000' } },
				],
			},
			'a:lin': { '@_ang': '0', '@_scaled': '1' },
			'a:tileRect': {},
		});
	});
});

describe('supportsChartGradientFill', () => {
	it('accepts area-filled types and rejects line-drawn and ChartEx ones', () => {
		for (const type of [
			'bar',
			'bar3D',
			'area',
			'pie',
			'doughnut',
			'ofPie',
			'bubble',
			'surface',
		] as const) {
			expect(supportsChartGradientFill(type)).toBeTruthy();
		}
		for (const type of [
			'line',
			'line3D',
			'scatter',
			'radar',
			'stock',
			'waterfall',
			'unknown',
		] as const) {
			expect(supportsChartGradientFill(type)).toBeFalsy();
		}
	});
});
