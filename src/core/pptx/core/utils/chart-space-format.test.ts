import { describe, expect, it } from 'vitest';

import type { PptxChartData, XmlObject } from '../types';
import { parseChartAreaFormat } from './chart-area-format';
import { applyChartAxisDeletedToXml } from './chart-axis-deleted';
import { applyChartSpaceFormatToXml } from './chart-space-format';

const localName = (key: string) => key.replace(/^.*:/u, '');
const THEME_WHITE = { 'a:schemeClr': { '@_val': 'bg1' } };
const resolveColor = (node: XmlObject) => (node['a:schemeClr'] ? '#FFFFFF' : undefined);

function data(overrides: Partial<PptxChartData>): PptxChartData {
	return { chartType: 'bar', categories: [], series: [], ...overrides };
}

function authoredSpace(): XmlObject {
	return {
		'@_xmlns:c': 'c',
		'c:date1904': { '@_val': '0' },
		'mc:AlternateContent': {},
		'c:chart': { 'c:plotArea': {} },
		'c:spPr': {
			'a:solidFill': THEME_WHITE,
			'a:ln': { '@_w': '9525', 'a:noFill': {} },
			'a:effectLst': {},
		},
		'c:txPr': {},
	};
}

describe('applyChartSpaceFormatToXml (loaded charts)', () => {
	it('leaves a chart the model still describes untouched', () => {
		const space = authoredSpace();
		const spPr = space['c:spPr'];
		const model = data({ style: { chartAreaFill: '#ffffff', chartAreaBorder: 'none' } });
		expect(applyChartSpaceFormatToXml(space, undefined, model, localName, { resolveColor })).toBe(
			false,
		);
		expect(space['c:spPr']).toBe(spPr);
		expect(space['c:roundedCorners']).toBeUndefined();
	});

	it('changes only the border, keeping the themed fill and the effects', () => {
		const space = authoredSpace();
		applyChartSpaceFormatToXml(
			space,
			undefined,
			data({ style: { chartAreaFill: '#FFFFFF', chartAreaBorder: '#FF0000' } }),
			localName,
			{ resolveColor },
		);
		expect(space['c:spPr']).toStrictEqual({
			'a:solidFill': THEME_WHITE,
			'a:ln': { '@_w': '9525', 'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } } },
			'a:effectLst': {},
		});
	});

	it('replaces the fill choice before a:ln', () => {
		const space = authoredSpace();
		applyChartSpaceFormatToXml(
			space,
			undefined,
			data({ style: { chartAreaFill: 'none', chartAreaBorder: 'none' } }),
			localName,
			{ resolveColor },
		);
		expect(Object.keys(space['c:spPr'] as XmlObject)).toStrictEqual([
			'a:noFill',
			'a:ln',
			'a:effectLst',
		]);
	});

	it('inserts roundedCorners before AlternateContent and a new spPr after c:chart', () => {
		const space: XmlObject = {
			'c:date1904': { '@_val': '0' },
			'mc:AlternateContent': {},
			'c:chart': {},
			'c:txPr': {},
		};
		applyChartSpaceFormatToXml(
			space,
			undefined,
			data({ roundedCorners: false, style: { chartAreaFill: 'none' } }),
			localName,
			{ resolveColor },
		);
		expect(Object.keys(space)).toStrictEqual([
			'c:date1904',
			'c:roundedCorners',
			'mc:AlternateContent',
			'c:chart',
			'c:spPr',
			'c:txPr',
		]);
	});

	it('removes a cleared roundedCorners and reconciles the plot area before extLst', () => {
		const space: XmlObject = { 'c:roundedCorners': { '@_val': '0' }, 'c:chart': {} };
		const plotArea: XmlObject = { 'c:layout': {}, 'c:barChart': {}, 'c:extLst': {} };
		applyChartSpaceFormatToXml(
			space,
			plotArea,
			data({ style: { plotAreaFill: 'none' } }),
			localName,
			{ resolveColor },
		);
		expect(space['c:roundedCorners']).toBeUndefined();
		expect(Object.keys(plotArea)).toStrictEqual(['c:layout', 'c:barChart', 'c:spPr', 'c:extLst']);
	});

	it('parses fill and border symmetrically', () => {
		expect(parseChartAreaFormat(authoredSpace(), localName, resolveColor)).toStrictEqual({
			fill: '#FFFFFF',
			border: 'none',
		});
	});
});

describe('applyChartAxisDeletedToXml', () => {
	it('writes only an explicit change, after c:scaling', () => {
		const axis: XmlObject = { 'c:axId': {}, 'c:scaling': {}, 'c:axPos': {} };
		expect(applyChartAxisDeletedToXml(axis, undefined, localName)).toBe(false);
		expect(applyChartAxisDeletedToXml(axis, false, localName)).toBe(false);
		applyChartAxisDeletedToXml(axis, true, localName);
		expect(Object.keys(axis)).toStrictEqual(['c:axId', 'c:scaling', 'c:delete', 'c:axPos']);
		expect(axis['c:delete']).toStrictEqual({ '@_val': '1' });
	});
});
