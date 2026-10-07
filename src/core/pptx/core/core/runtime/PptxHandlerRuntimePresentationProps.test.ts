/**
 * @fileoverview Tests for `extractChartStyle`'s legend-position resolution.
 *
 * `extractChartStyle` is protected on a deeply mixed-in class, so it is
 * exercised via a `this`-shaped stub carrying just `xmlLookupService`/
 * `compatibilityService`/`parseColor`, and hand-crafted XML object trees
 * mirroring what `fast-xml-parser` produces (see
 * `PptxHandlerRuntimeChartChrome.test.ts` for the same pattern).
 *
 * Importing `PptxHandlerRuntimePresentationProps` directly (rather than a
 * file further along the mixin chain) trips a circular-import ordering
 * problem in this module graph ("Class extends value undefined").
 * `extractChartStyle` is still reachable, inherited, off
 * `PptxHandlerRuntimeSaveDataSerialization` (PresentationProps ->
 * SaveSlideUtils -> ... -> SaveDataSerialization), which loads safely; using
 * `Object.create` on its prototype (rather than a bare object stub) also
 * picks up the private helpers `extractChartStyle` calls internally
 * (`parseChartContainerFill`) for free.
 */

import { describe, it, expect } from 'vitest';

import { PptxXmlLookupService } from '../../services/PptxXmlLookupService';
import type { PptxChartStyle, XmlObject } from '../../types';
// Importing ChartParsingHelpers (lower in the mixin chain) first primes module
// evaluation order so the SaveDataSerialization import below doesn't trip the
// circular-import ordering hazard (see PptxHandlerRuntimeChartChrome.test.ts,
// which loads both for the same reason).
import './PptxHandlerRuntimeChartParsingHelpers';
import { PptxHandlerRuntime } from './PptxHandlerRuntimeSaveDataSerialization';

const xmlLookupService = new PptxXmlLookupService();

function getLocalName(qualifiedName: string): string {
	const colonIndex = qualifiedName.lastIndexOf(':');
	return colonIndex >= 0 ? qualifiedName.substring(colonIndex + 1) : qualifiedName;
}

const compatibilityService = { getXmlLocalName: getLocalName };

interface ExtractChartStyleHost {
	extractChartStyle(
		chartSpace: XmlObject | undefined,
		chartRoot: XmlObject | undefined,
	): PptxChartStyle | undefined;
}

function extractChartStyle(
	chartSpace: XmlObject | undefined,
	chartRoot: XmlObject | undefined,
	themeFontMap: Record<string, string> = {},
): PptxChartStyle | undefined {
	const instance = Object.create(PptxHandlerRuntime.prototype) as ExtractChartStyleHost &
		Record<string, unknown>;
	instance.xmlLookupService = xmlLookupService;
	instance.compatibilityService = compatibilityService;
	instance.parseColor = () => undefined;
	instance.themeFontMap = themeFontMap;
	return instance.extractChartStyle(chartSpace, chartRoot);
}

describe('extractChartStyle legend position (C2-G5)', () => {
	it('reads classic c:legend/c:legendPos/@val (child-element form)', () => {
		const chartRoot: XmlObject = {
			'c:legend': {
				'c:legendPos': { '@_val': 'r' },
			},
		};
		const style = extractChartStyle(undefined, chartRoot);
		expect(style?.legendPosition).toBe('r');
	});

	it('falls back to cx:legend/@pos (attribute form) when no child c:legendPos exists', () => {
		// ChartEx (cx:) charts carry the position directly on the element
		// instead of nesting a legendPos child, per CT_Legend (chartex schema).
		const chartRoot: XmlObject = {
			'cx:legend': { '@_pos': 't' },
		};
		const style = extractChartStyle(undefined, chartRoot);
		expect(style?.legendPosition).toBe('t');
	});

	it('prefers the child element over the attribute when both are somehow present', () => {
		const chartRoot: XmlObject = {
			'c:legend': {
				'@_pos': 'l',
				'c:legendPos': { '@_val': 'r' },
			},
		};
		const style = extractChartStyle(undefined, chartRoot);
		expect(style?.legendPosition).toBe('r');
	});

	it('leaves a missing c:legendPos unset so a save does not add one', () => {
		const chartRoot: XmlObject = { 'c:legend': { 'c:overlay': { '@_val': '0' } } };
		const style = extractChartStyle(undefined, chartRoot);
		expect(style?.hasLegend).toBe(true);
		expect(style?.legendPosition).toBeUndefined();
	});
});

// Measured with PowerPoint 16 (Legend.IncludeInLayout): c:overlay alone decides
// whether the legend reserves space, for every position including `tr`, and a
// classic legend with no c:overlay element floats over the plot.
describe('extractChartStyle legend overlay', () => {
	const legendWith = (children: XmlObject): XmlObject => ({
		'c:legend': { 'c:legendPos': { '@_val': 'b' }, ...children },
	});

	it('reserves space for c:overlay val="0"', () => {
		const style = extractChartStyle(undefined, legendWith({ 'c:overlay': { '@_val': '0' } }));
		expect(style?.legendOverlay).toBe(false);
	});

	it('overlays for c:overlay val="1" and for a bare <c:overlay/>', () => {
		expect(
			extractChartStyle(undefined, legendWith({ 'c:overlay': { '@_val': '1' } }))?.legendOverlay,
		).toBe(true);
		expect(extractChartStyle(undefined, legendWith({ 'c:overlay': {} }))?.legendOverlay).toBe(true);
	});

	it('overlays a classic legend with no c:overlay element', () => {
		expect(extractChartStyle(undefined, legendWith({}))?.legendOverlay).toBe(true);
	});

	it('reads cx:legend/@overlay and leaves an unstated ChartEx overlay unknown', () => {
		const chartSpace: XmlObject = { 'cx:chartData': {} };
		expect(
			extractChartStyle(chartSpace, { 'cx:legend': { '@_pos': 't', '@_overlay': '1' } })
				?.legendOverlay,
		).toBe(true);
		expect(
			extractChartStyle(chartSpace, { 'cx:legend': { '@_pos': 't' } })?.legendOverlay,
		).toBeUndefined();
	});
});

// Regression: a chart authoring an 18pt legend/chart-level default text style
// (`c:legend/c:txPr` or `c:chartSpace/c:txPr`) rendered its legend at a fixed
// 9px default, because only PER-ENTRY `c:legendEntry/c:txPr` overrides were
// ever read. See `chart-legend-entries.ts`'s `applyLegendEntryOverrides`.
describe('extractChartStyle legend text style (chart-level c:txPr default)', () => {
	function defRPrTxPr(sz: number): XmlObject {
		return {
			'a:p': { 'a:pPr': { 'a:defRPr': { '@_sz': String(sz) } } },
		};
	}

	it("reads the legend's own c:txPr as the legend text default", () => {
		const chartRoot: XmlObject = {
			'c:legend': {
				'c:legendPos': { '@_val': 'b' },
				'c:txPr': defRPrTxPr(1800),
			},
		};
		const style = extractChartStyle(undefined, chartRoot);
		expect(style?.legendTextStyle?.fontSize).toBe(18);
	});

	it('falls back to c:chartSpace/c:txPr when the legend has none of its own', () => {
		const chartSpace: XmlObject = { 'c:txPr': defRPrTxPr(1800) };
		const chartRoot: XmlObject = { 'c:legend': { 'c:legendPos': { '@_val': 'b' } } };
		const style = extractChartStyle(chartSpace, chartRoot);
		expect(style?.legendTextStyle?.fontSize).toBe(18);
	});

	it("prefers the legend's own c:txPr over the chart-space default", () => {
		const chartSpace: XmlObject = { 'c:txPr': defRPrTxPr(1800) };
		const chartRoot: XmlObject = {
			'c:legend': { 'c:legendPos': { '@_val': 'b' }, 'c:txPr': defRPrTxPr(1000) },
		};
		const style = extractChartStyle(chartSpace, chartRoot);
		expect(style?.legendTextStyle?.fontSize).toBe(10);
	});

	it('leaves legendTextStyle undefined when neither the legend nor the chart space has a txPr', () => {
		const chartRoot: XmlObject = { 'c:legend': { 'c:legendPos': { '@_val': 'b' } } };
		const style = extractChartStyle(undefined, chartRoot);
		expect(style?.legendTextStyle).toBeUndefined();
	});
});

describe('extractChartStyle chart-wide text style', () => {
	it('reads c:chartSpace/c:txPr, East Asian face included', () => {
		const chartSpace: XmlObject = {
			'c:txPr': {
				'a:bodyPr': {},
				'a:p': {
					'a:pPr': {
						'a:defRPr': {
							'@_sz': '900',
							'a:latin': { '@_typeface': 'Arial' },
							'a:ea': { '@_typeface': 'Malgun Gothic' },
						},
					},
				},
			},
		};
		expect(extractChartStyle(chartSpace, {})?.textStyle).toStrictEqual({
			fontSize: 9,
			fontFamily: 'Arial',
			eastAsiaFontFamily: 'Malgun Gothic',
		});
	});

	it('leaves it unset when the chart has no c:txPr and the theme no fonts', () => {
		expect(extractChartStyle({}, { 'c:legend': {} })?.textStyle).toBeUndefined();
	});

	it("starts from the theme's minor fonts", () => {
		const theme = { 'mn-lt': 'Calibri', 'mn-ea': 'Malgun Gothic' };
		expect(extractChartStyle({}, {}, theme)?.textStyle).toStrictEqual({
			fontFamily: 'Calibri',
			eastAsiaFontFamily: 'Malgun Gothic',
		});
		const chartSpace: XmlObject = {
			'c:txPr': { 'a:p': { 'a:pPr': { 'a:defRPr': { 'a:latin': { '@_typeface': 'Arial' } } } } },
		};
		expect(extractChartStyle(chartSpace, {}, theme)?.textStyle).toStrictEqual({
			fontFamily: 'Arial',
			eastAsiaFontFamily: 'Malgun Gothic',
		});
	});
});
