import { PptxRuntimeDependencyFactory } from '../core/factories/PptxRuntimeDependencyFactory';
import { parseXml } from '../../../xml/index';
import { shapeProperties } from '../../../chart/parse-util';
import type { PptxChartType } from '../types';
/**
 * Series colour extraction for area-fill vs line-drawn chart families (moved from
 * `PptxHandlerRuntimeChartSeriesColor.test.ts` with the reader it pins).
 *
 * A series' explicit colour lives at `c:ser/c:spPr/a:solidFill` for area-filled types, but LINE
 * (and scatter/radar/stock) series author it on the outline: `c:ser/c:spPr/a:ln/a:solidFill`. The
 * rule is keyed off the series' own group type, always resolved, never off the combo-only
 * `seriesChartType` tag. The full pipeline (plain and combo charts) is pinned in
 * `chart-from-neutral.test.ts`.
 */
import { describe, expect, it } from 'vitest';

import { PptxXmlLookupService } from '../services/PptxXmlLookupService';
import type { XmlObject } from '../types';
import {
	readSeriesTreeFields as readFields,
	type ChartTreeReader,
} from './chart-series-tree-fields';

/** Minimal `parseColor`: resolve a solidFill node's `a:srgbClr/@val`. */
function parseColorStub(node: XmlObject | undefined): string | undefined {
	const srgb = node?.['a:srgbClr'] as XmlObject | undefined;
	const val = srgb?.['@_val'];
	return typeof val === 'string' && val.length > 0 ? `#${val}` : undefined;
}

const STUB_READER: ChartTreeReader = {
	xml: new PptxXmlLookupService(),
	parseColor: parseColorStub,
	resolveTypeface: (raw) => raw,
	localName: (key) => key.slice(key.lastIndexOf(':') + 1),
};

function readSeriesTreeFields(node: XmlObject, type: PptxChartType, reader: ChartTreeReader) {
	const body = new PptxRuntimeDependencyFactory().createBuilder().build(node);
	const root = parseXml(
		'<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
			body +
			'</c:ser>',
	).documentElement;
	return readFields(node, type, reader, { spPr: shapeProperties(root), dataPoints: [] });
}

describe('readSeriesTreeFields colour extraction', () => {
	it('reads the direct spPr solidFill (bar/area style)', () => {
		const fields = readSeriesTreeFields(
			{ 'c:spPr': { 'a:solidFill': { 'a:srgbClr': { '@_val': '4472C4' } } } },
			'bar',
			STUB_READER,
		);
		expect(fields.color).toBe('#4472C4');
	});

	it('falls back to the a:ln solidFill for a LINE series', () => {
		const fields = readSeriesTreeFields(
			{
				'c:spPr': {
					'a:ln': { '@_w': '28575', 'a:solidFill': { 'a:srgbClr': { '@_val': 'ED7D31' } } },
				},
			},
			'line',
			STUB_READER,
		);
		expect(fields.color).toBe('#ED7D31');
		expect(fields.lineWidth).toBe(2.25);
	});

	it('does NOT take an outline colour as the fill of a bar series', () => {
		const fields = readSeriesTreeFields(
			{ 'c:spPr': { 'a:ln': { 'a:solidFill': { 'a:srgbClr': { '@_val': 'ED7D31' } } } } },
			'bar',
			STUB_READER,
		);
		expect(fields.color).toBeUndefined();
	});

	it('the direct fill wins over the outline for a line series carrying both', () => {
		const fields = readSeriesTreeFields(
			{
				'c:spPr': {
					'a:solidFill': { 'a:srgbClr': { '@_val': '70AD47' } },
					'a:ln': { 'a:solidFill': { 'a:srgbClr': { '@_val': 'ED7D31' } } },
				},
			},
			'line',
			STUB_READER,
		);
		expect(fields.color).toBe('#70AD47');
	});

	it('leaves the colour undefined for a scatter series whose line is noFill', () => {
		const fields = readSeriesTreeFields(
			{ 'c:spPr': { 'a:ln': { 'a:noFill': {} } } },
			'scatter',
			STUB_READER,
		);
		expect(fields.color).toBeUndefined();
		expect(fields.lineNoFill).toBe(true);
	});

	it('reads a line series outline gradient (a:ln/a:gradFill) as lineGradientFill', () => {
		const codec = {
			extractGradientStops: () => [
				{ color: '#C08FBF', position: 0 },
				{ color: '#800F85', position: 100 },
			],
			extractGradientType: () => 'linear' as const,
			extractGradientAngle: () => 0,
			extractGradientFocalPoint: () => undefined,
		};
		const fields = readSeriesTreeFields(
			{
				'c:spPr': {
					'a:ln': {
						'@_w': '76200',
						'a:gradFill': {
							'a:gsLst': {
								'a:gs': [
									{ '@_pos': '0', 'a:srgbClr': { '@_val': 'C08FBF' } },
									{ '@_pos': '100000', 'a:srgbClr': { '@_val': '800F85' } },
								],
							},
							'a:lin': { '@_ang': '0' },
						},
					},
				},
			},
			'line',
			{ ...STUB_READER, colorStyleCodec: codec },
		);
		expect(fields.lineGradientFill).toStrictEqual({
			type: 'linear',
			angle: 0,
			stops: [
				{ color: '#C08FBF', position: 0 },
				{ color: '#800F85', position: 100 },
			],
		});
		expect(fields.gradientFill).toBeUndefined();
		expect(fields.lineWidth).toBe(6);
	});
});

it('uses neutral paint values even when a compatibility source tree contains stale values', () => {
	const fields = readFields(
		{
			'c:spPr': {
				'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } },
				'a:ln': { '@_w': '12700' },
			},
		},
		'line',
		{
			...STUB_READER,
			parseColor: (node) =>
				(node?.['a:schemeClr'] as XmlObject | undefined)?.['@_val'] === 'accent1'
					? '#204060'
					: undefined,
		},
		{
			spPr: {
				fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1', transforms: [] } },
				line: { widthEmu: 38100, dash: 'dash', fill: { kind: 'none' } },
			},
			dataPoints: [],
		},
	);
	expect(fields).toMatchObject({
		color: '#204060',
		lineWidth: 3,
		lineDashStyle: 'dash',
		lineNoFill: true,
	});
});

it('projects neutral marker and point overrides while retaining only picture and identity source facts', () => {
	const fields = readFields({}, 'bar', STUB_READER, {
		marker: { symbol: 'diamond', size: 8, spPr: { line: { fill: { kind: 'none' } } } },
		dataPoints: [
			{
				index: 3,
				explosion: 12,
				bubble3D: false,
				spPr: { fill: { kind: 'solid', color: { kind: 'srgb', value: '123456', transforms: [] } } },
				marker: { symbol: 'circle', size: 4 },
			},
		],
	});
	expect(fields.marker).toEqual({ symbol: 'diamond', size: 8, spPr: { lineNoFill: true } });
	expect(fields.dataPoints).toEqual([
		{
			idx: 3,
			spPr: { fillColor: '#123456' },
			explosion: 12,
			marker: { symbol: 'circle', size: 4 },
			bubble3D: false,
		},
	]);
});
