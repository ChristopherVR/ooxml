import { describe, expect, it } from 'vitest';
import { NS, elements, first, parseXml } from '../../xml/index';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { parseDrawingEffects } from '../../drawingml/drawing-effects';
import { parseDrawingFill } from '../../drawingml/drawing-fill';
import { setChartTrendline, setChartErrorBars } from './analysis';
import { setChartMarker } from './marker';
import { setChartTargetEffect } from './effects';
import { chartTargetTextFormatting, setChartTargetText } from './text';
import { editChartTextBody } from './text';
import { chartTargetThreeD, setChartTargetThreeD, setChartView3D } from './three-d';
import { setChartSeriesSecondaryAxis } from './secondary-axis';
import { setChartType } from './type';
import { applyChartEditIntent } from './intent';
import { setChartTargetLineOptions } from './line-options';
import { setChartTargetFill } from './target-paint';
import { columnChart } from './test-support/column-chart';

const color = {
	kind: 'scheme' as const,
	value: 'accent2',
	transforms: [{ name: 'alpha', value: '40000' }],
};
const reopen = (space: ReturnType<typeof columnChart>) => {
	const result = parseChartSpace(writeChartSpace(space));
	expect(result.issues).toEqual([]);
	return result.chartSpace;
};

describe('shared advanced classic chart operations', () => {
	it('edits standalone and East Asian text through the same validated source-preserving helper', () => {
		const source = columnChart().title!.tx!.rich!;
		const before = structuredClone(source);
		const body = editChartTextBody(source, { eastAsiaTypeface: 'Yu Gothic', sizePt: 20 });
		expect(body.sourceXml).toContain('<a:ea typeface="Yu Gothic"');
		expect(body.text).toBe(source.text);
		expect(source).toEqual(before);
		expect(editChartTextBody(body, { eastAsiaTypeface: null }).sourceXml).not.toContain('<a:ea');
		expect(() => editChartTextBody(body, { sizePt: -1 })).toThrow('text size');
		const updated = setChartTargetText(
			columnChart(),
			{ kind: 'title' },
			{ eastAsiaTypeface: 'Yu Gothic' },
		);
		expect(chartTargetTextFormatting(updated, { kind: 'title' }).eastAsiaTypeface).toBe(
			'Yu Gothic',
		);
	});
	it('clears an outline width and edits arrow kind without losing arrow dimensions or fill', () => {
		const source = columnChart();
		source.spPr = {
			line: {
				widthEmu: 12700,
				compound: 'dbl',
				headEnd: { type: 'triangle', width: 'lg', length: 'sm' },
				fill: { kind: 'solid', color },
			},
		};
		const imported = reopen(source);
		const edited = setChartTargetLineOptions(
			imported,
			{ kind: 'chartArea' },
			{ widthEmu: null, headEnd: { type: 'oval' } },
		);
		const saved = reopen(edited).spPr!.line!;
		expect(saved.widthEmu).toBeUndefined();
		expect(saved).toMatchObject({
			compound: 'dbl',
			headEnd: { type: 'oval', width: 'lg', length: 'sm' },
			fill: { color },
		});
	});
	it('uses the shared pattern writer through the actual chart save consumer', () => {
		const source = columnChart();
		source.spPr = {
			sourceXml: `<a:pattFill xmlns:v="urn:v" prst="pct5" v:keep="yes"><a:fgClr><a:schemeClr val="accent1"><v:color/></a:schemeClr></a:fgClr><a:extLst><a:ext uri="keep"><v:data/></a:ext></a:extLst></a:pattFill>`,
		};
		source.spPr.fill = parseDrawingFill(
			parseXml(`<c:spPr xmlns:c="${NS.c}" xmlns:a="${NS.a}">${source.spPr.sourceXml}</c:spPr>`)
				.documentElement,
		)!;
		const imported = reopen(source);
		const fill = imported.spPr!.fill!;
		if (fill.kind !== 'pattern') throw new Error('Missing pattern');
		const edited = setChartTargetFill(
			imported,
			{ kind: 'chartArea' },
			{ ...fill, preset: 'pct10' },
		);
		const saved = writeChartSpace(edited);
		expect(saved).toContain('v:keep="yes"');
		expect(saved).toContain('<v:color');
		expect(saved).toContain('<v:data');
		expect(parseChartSpace(saved).chartSpace.spPr!.fill).toMatchObject({
			kind: 'pattern',
			preset: 'pct10',
		});
	});
	it('adds, edits and removes independent trendlines while retaining other series and source cells', () => {
		const source = columnChart();
		const before = structuredClone(source);
		const linear = setChartTrendline(source, 0, 0, { type: 'linear', displayEquation: true });
		const polynomial = setChartTrendline(linear, 0, 1, { type: 'poly', order: 4, forward: 2 });
		const saved = reopen(polynomial);
		expect(saved.plotArea.groups[0]!.series[0]!.trendlines).toMatchObject([
			{ trendlineType: 'linear', displayEquation: true },
			{ trendlineType: 'poly', order: 4, forward: 2 },
		]);
		expect(source).toEqual(before);
		expect(saved.plotArea.groups[0]!.series[0]!.values).toEqual(
			source.plotArea.groups[0]!.series[0]!.values,
		);
		expect(
			setChartTrendline(polynomial, 0, 0, null).plotArea.groups[0]!.series[0]!.trendlines,
		).toHaveLength(1);
		expect(() => setChartTrendline(source, 0, 0, { type: 'poly', order: 1 })).toThrow();
		expect(() => setChartTrendline(setChartType(source, { kind: 'pie' }), 0, 0, {})).toThrow();
	});
	it('writes ordinary and custom error bars and refuses incompatible directions and sources', () => {
		const source = columnChart();
		const fixed = setChartErrorBars(source, 0, 'y', {
			valueType: 'percentage',
			value: 5,
			noEndCap: true,
		});
		expect(reopen(fixed).plotArea.groups[0]!.series[0]!.errorBars?.[0]).toMatchObject({
			direction: 'y',
			valueType: 'percentage',
			value: 5,
			noEndCap: true,
		});
		const amounts = {
			kind: 'numLit' as const,
			cache: { type: 'number' as const, points: [{ index: 0, value: '2' }] },
		};
		const custom = setChartErrorBars(fixed, 0, 'y', {
			valueType: 'cust',
			plus: amounts,
			minus: amounts,
		});
		expect(
			reopen(custom).plotArea.groups[0]!.series[0]!.errorBars?.[0]?.plus?.cache?.points[0]?.value,
		).toBe('2');
		expect(() => setChartErrorBars(source, 0, 'x', {})).toThrow();
		expect(() => setChartErrorBars(source, 0, 'y', { valueType: 'cust' })).toThrow();
		expect(() => setChartErrorBars(source, 0, 'y', { value: -1 })).toThrow();
		expect(
			setChartErrorBars(custom, 0, 'y', null).plotArea.groups[0]!.series[0]!.errorBars,
		).toEqual([]);
	});
	it('retains marker styles on partial edits and disables markers explicitly', () => {
		const line = setChartType(columnChart(), { kind: 'line' });
		const styled = setChartMarker(line, 0, {
			symbol: 'diamond',
			size: 9,
			spPr: { fill: { kind: 'solid', color } },
		});
		const resized = setChartMarker(styled, 0, { size: 12 });
		expect(reopen(resized).plotArea.groups[0]!.series[0]!.marker).toMatchObject({
			symbol: 'diamond',
			size: 12,
			spPr: { fill: { color } },
		});
		expect(setChartMarker(resized, 0, null).plotArea.groups[0]!.series[0]!.marker?.symbol).toBe(
			'none',
		);
		expect(() => setChartMarker(columnChart(), 0, { size: 9 })).toThrow();
		expect(() => setChartMarker(line, 0, { size: 73 })).toThrow();
	});
	it('updates one shape effect while retaining unknown effects, transforms and authored tokens', () => {
		const source = columnChart();
		source.spPr = {
			effectsXml: `<a:effectLst xmlns:a="${NS.a}" xmlns:v="urn:v"><a:glow rad="12700"><a:schemeClr val="accent3"><v:keep/></a:schemeClr></a:glow><v:unknown/></a:effectLst>`,
		};
		const edited = setChartTargetEffect(
			source,
			{ kind: 'chartArea' },
			{
				kind: 'outerShadow',
				blurRadiusEmu: 12700,
				distanceEmu: 25400,
				directionAngle: 5400000,
				color,
			},
		);
		const saved = reopen(edited).spPr!.effectsXml!;
		expect(saved).toContain('<v:unknown');
		expect(saved).toContain('<v:keep');
		expect(parseDrawingEffects(parseXml(saved).documentElement)?.effects).toMatchObject([
			{ kind: 'glow' },
			{ kind: 'unsupported' },
			{ kind: 'outerShadow', color },
		]);
		const cleared = setChartTargetEffect(
			edited,
			{ kind: 'chartArea' },
			{ kind: 'outerShadow', remove: true },
		);
		expect(cleared.spPr!.effectsXml).toContain('<v:unknown');
		expect(cleared.spPr!.effectsXml).not.toContain('outerShdw');
	});
	it('patches rich title runs, paragraph defaults and textbox geometry without losing text', () => {
		const source = columnChart();
		const title = source.title!.text;
		const edited = setChartTargetText(
			source,
			{ kind: 'title' },
			{
				fill: { kind: 'solid', color },
				line: { widthEmu: 12700, fill: { kind: 'none' } },
				rotation: 20,
				direction: 'vert',
				anchor: 't',
				sizePt: 18,
				typeface: 'Arial',
			},
		);
		const saved = reopen(edited);
		expect(saved.title!.text).toBe(title);
		expect(chartTargetTextFormatting(saved, { kind: 'title' })).toMatchObject({
			rotation: 20,
			direction: 'vert',
			anchor: 't',
			sizePt: 18,
			typeface: 'Arial',
			fill: { color },
		});
		const rich = parseXml(
			`<c:rich xmlns:c="${NS.c}" xmlns:a="${NS.a}">${saved.title!.tx!.rich!.sourceXml}</c:rich>`,
		).documentElement;
		const props = first(first(rich, 'p', NS.a), 'r', NS.a)!;
		expect(elements(first(props, 'rPr', NS.a)!).map((node) => node.localName)).toEqual([
			'ln',
			'solidFill',
			'latin',
		]);
		expect(() =>
			setChartTargetText(source, { kind: 'series', seriesIndex: 0 }, { sizePt: 12 }),
		).toThrow();
	});
	it('writes source-backed bevels, depth, contour and lighting without clearing existing shape paint', () => {
		const source = columnChart();
		const original = source.plotArea.groups[0]!.series[0]!.spPr!.fill;
		const edited = setChartTargetThreeD(
			source,
			{ kind: 'series', seriesIndex: 0 },
			{
				bevelTop: { preset: 'circle', widthEmu: 25400, heightEmu: 12700 },
				depthEmu: 38100,
				depthColor: color,
				material: 'plastic',
				lighting: 'balanced',
			},
		);
		const saved = reopen(edited);
		expect(saved.plotArea.groups[0]!.series[0]!.spPr!.fill).toEqual(original);
		expect(chartTargetThreeD(saved, { kind: 'series', seriesIndex: 0 })).toMatchObject({
			bevelTop: { preset: 'circle', widthEmu: 25400 },
			depthEmu: 38100,
			depthColor: color,
			material: 'plastic',
			lighting: 'balanced',
		});
		expect(() => setChartView3D(source, { rotX: 30 })).toThrow();
		const three = setChartType(source, { kind: 'bar', is3D: true });
		expect(
			reopen(setChartView3D(three, { rotX: 30, rotY: 45, perspective: 20 })).view3D,
		).toMatchObject({ rotX: 30, rotY: 45, perspective: 20 });
	});
	it('moves a coherent combo series between axes while retaining its data and document-order target', () => {
		const source = columnChart();
		const group = source.plotArea.groups[0]!;
		const series = group.series.pop()!;
		source.plotArea.groups.push({
			kind: 'line',
			element: 'lineChart',
			is3D: false,
			series: [series],
			axisIds: [...group.axisIds],
		});
		const edited = setChartSeriesSecondaryAxis(source, 1, true);
		const saved = reopen(edited);
		expect(saved.plotArea.groups[1]!.series[0]!.values).toEqual(series.values);
		expect(saved.plotArea.groups[1]!.axisIds).not.toEqual(group.axisIds);
		expect(saved.plotArea.axes).toHaveLength(4);
		expect(setChartSeriesSecondaryAxis(saved, 1, false).plotArea.groups[1]!.axisIds).toEqual(
			group.axisIds,
		);
		expect(() => setChartSeriesSecondaryAxis(columnChart(), 0, true)).toThrow();
	});
	it('routes new type and analysis intent variants through the same pure operations', () => {
		const source = columnChart();
		const line = applyChartEditIntent(source, { op: 'type', type: { kind: 'line' } });
		expect(
			applyChartEditIntent(line, {
				op: 'marker',
				seriesIndex: 0,
				edit: { symbol: 'circle', size: 7 },
			}).plotArea.groups[0]!.series[0]!.marker?.size,
		).toBe(7);
		expect(
			applyChartEditIntent(source, {
				op: 'trendline',
				seriesIndex: 0,
				index: 0,
				edit: { type: 'linear' },
			}).plotArea.groups[0]!.series[0]!.trendlines,
		).toHaveLength(1);
	});
});
