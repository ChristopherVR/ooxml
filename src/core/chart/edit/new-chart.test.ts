import { describe, expect, it } from 'vitest';

import type { ChartDataSheet } from '../data-sheet';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { newChartSpace, type NewChartOptions } from './new-chart';
import type { ChartTypeTarget } from './type';

const SHEET: ChartDataSheet = {
	categories: ['North', 'South'],
	series: [
		{ name: 'Revenue', values: [4, null] },
		{ name: 'Cost', values: [2.5, 3] },
		{ name: 'Margin', values: [1, 0] },
	],
};

const MINIMAL: NewChartOptions = { data: 'literal', profile: 'minimal' };

const xmlOf = (target?: ChartTypeTarget, options: NewChartOptions = MINIMAL) =>
	writeChartSpace(newChartSpace(SHEET, target, options));

/** Writes, reads back and checks that the reader reports nothing it could not hold. */
function reread(target?: ChartTypeTarget, options: NewChartOptions = MINIMAL) {
	const parsed = parseChartSpace(xmlOf(target, options));
	expect(parsed.issues).toStrictEqual([]);
	return parsed.chartSpace;
}

describe('newChartSpace: Office defaults', () => {
	it('keeps the workbook formulas, Excel axes, gap, overlap and bottom legend', () => {
		const xml = writeChartSpace(newChartSpace(SHEET));
		expect(xml).toContain('<c:roundedCorners val="0"/>');
		expect(xml).toContain('<c:f>Sheet1!$B$1</c:f>');
		expect(xml).toContain('<c:gapWidth val="219"/><c:overlap val="-27"/>');
		expect(xml).toContain('<c:majorGridlines/>');
		expect(xml).toContain('<c:legendPos val="b"/>');
		expect(xml).toContain('<c:axId val="500000001"/><c:axId val="500000002"/>');
	});

	it('renumbers the axes, crossings included, when ids are given', () => {
		const space = newChartSpace(SHEET, { kind: 'line', is3D: true }, { axisIds: [7, 8, 9] });
		expect(space.plotArea.groups[0]!.axisIds).toStrictEqual([7, 8, 9]);
		expect(space.plotArea.axes.map((axis) => [axis.id, axis.crossAxisId])).toStrictEqual([
			[7, 8],
			[8, 7],
			[9, 8],
		]);
	});
});

describe('newChartSpace: literal data', () => {
	it('writes names as c:v and data as literals, leaving blank points out of the count', () => {
		const xml = xmlOf();
		expect(xml).not.toContain('<c:f>');
		expect(xml).toContain('<c:tx><c:v>Revenue</c:v></c:tx>');
		expect(xml).toContain(
			'<c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strLit></c:cat>',
		);
		expect(xml).toContain(
			'<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val>',
		);
		const series = reread().plotArea.groups[0]!.series;
		expect(series.map((item) => item.tx?.text)).toStrictEqual(['Revenue', 'Cost', 'Margin']);
		expect(series[2]!.values?.cache?.points.map((point) => point.value)).toStrictEqual(['1', '0']);
	});

	it('keeps the workbook-free data in the Office profile too', () => {
		const xml = xmlOf(undefined, { data: 'literal' });
		expect(xml).not.toContain('<c:f>');
		expect(xml).toContain('<c:invertIfNegative val="0"/>');
	});
});

describe('newChartSpace: the minimal profile', () => {
	it('writes a column chart with no gap, overlap, corners, gridlines or tick marks', () => {
		const xml = xmlOf();
		for (const absent of [
			'gapWidth',
			'overlap',
			'smooth',
			'roundedCorners',
			'majorGridlines',
			'majorTickMark',
			'invertIfNegative',
			'view3D',
		])
			expect(xml).not.toContain(`<c:${absent}`);
		expect(xml).toContain(
			'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>',
		);
		expect(xml).toContain(
			'<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="500000002"/></c:catAx>',
		);
	});

	it('writes what Office reads differently when absent: smoothing off, stacked bars overlapping', () => {
		const line = reread({ kind: 'line' }).plotArea.groups[0]!;
		expect(line.series.map((series) => series.smooth)).toStrictEqual([false, false, false]);
		expect(reread({ kind: 'stock' }).plotArea.groups[0]!.series[0]!.smooth).toBe(false);
		expect(reread({ kind: 'scatter' }).plotArea.groups[0]!.series[0]!.smooth).toBe(false);
		expect(
			reread({ kind: 'line', is3D: true }).plotArea.groups[0]!.series[0]!.smooth,
		).toBeUndefined();
		expect(reread({ kind: 'bar', grouping: 'stacked' }).plotArea.groups[0]!.overlap).toBe(100);
		expect(reread({ kind: 'bar', grouping: 'percentStacked' }).plotArea.groups[0]!.overlap).toBe(
			100,
		);
		expect(reread({ kind: 'bar' }).plotArea.groups[0]!.overlap).toBeUndefined();
		expect(
			reread({ kind: 'bar', grouping: 'stacked', is3D: true }).plotArea.groups[0]!.overlap,
		).toBeUndefined();
	});

	it('places the axes of horizontal bars at the left and the bottom', () => {
		const axes = reread({ kind: 'bar', barDirection: 'bar' }).plotArea.axes;
		expect(axes.map((axis) => [axis.kind, axis.position])).toStrictEqual([
			['cat', 'l'],
			['val', 'b'],
		]);
	});

	it('writes a clustered line or area as standard, the only lower grouping ST_Grouping has', () => {
		expect(xmlOf({ kind: 'line', grouping: 'clustered' })).toContain(
			'<c:grouping val="standard"/>',
		);
		expect(xmlOf({ kind: 'area', grouping: 'clustered' })).toContain(
			'<c:grouping val="standard"/>',
		);
		expect(xmlOf({ kind: 'line', grouping: 'stacked' })).toContain('<c:grouping val="stacked"/>');
		expect(xmlOf({ kind: 'bar', grouping: 'standard', is3D: true })).toContain(
			'<c:bar3DChart><c:barDir val="col"/><c:grouping val="standard"/>',
		);
	});

	it("writes each type's required settings and axes", () => {
		const group = (target: ChartTypeTarget) => reread(target).plotArea.groups[0]!;
		expect(group({ kind: 'pie' })).toMatchObject({ varyColors: true, axisIds: [] });
		expect(reread({ kind: 'doughnut' }).plotArea.axes).toStrictEqual([]);
		expect(group({ kind: 'ofPie' })).toMatchObject({ ofPieType: 'pie', varyColors: true });
		expect(group({ kind: 'scatter' })).toMatchObject({ scatterStyle: 'lineMarker' });
		const scatter = group({ kind: 'scatter' }).series[0]!;
		expect([scatter.xValues?.kind, scatter.yValues?.kind, scatter.categories]).toStrictEqual([
			'strLit',
			'numLit',
			undefined,
		]);
		expect(group({ kind: 'bubble' }).series[0]!.yValues?.kind).toBe('numLit');
		expect(reread({ kind: 'scatter' }).plotArea.axes.map((axis) => axis.kind)).toStrictEqual([
			'val',
			'val',
		]);
		expect(group({ kind: 'radar' })).toMatchObject({ radarStyle: 'marker', varyColors: false });
		expect(group({ kind: 'stock' }).varyColors).toBeUndefined();
		for (const target of [
			{ kind: 'line', is3D: true },
			{ kind: 'surface', is3D: false },
			{ kind: 'surface', is3D: true },
		] as const) {
			const space = reread(target);
			expect(space.plotArea.axes.map((axis) => axis.kind)).toStrictEqual(['cat', 'val', 'ser']);
			expect(space.plotArea.groups[0]!.axisIds).toHaveLength(3);
		}
	});

	it('refuses a stock chart without three or four price series', () => {
		const sheet = { categories: ['A'], series: [{ name: 'High', values: [1] }] };
		expect(() => newChartSpace(sheet, { kind: 'stock' }, MINIMAL)).toThrow(RangeError);
	});

	it('takes the axis ids, a series axis id included', () => {
		const space = reread({ kind: 'surface' }, { ...MINIMAL, axisIds: [11, 22, 33] });
		expect(space.plotArea.axes.map((axis) => [axis.id, axis.crossAxisId])).toStrictEqual([
			[11, 22],
			[22, 11],
			[33, 22],
		]);
	});

	it('refuses negative, fractional or repeated axis ids', () => {
		for (const axisIds of [
			[-1, 2],
			[1.5, 2],
			[3, 3],
		])
			expect(() => newChartSpace(SHEET, undefined, { ...MINIMAL, axisIds })).toThrow(RangeError);
	});
});

describe('newChartSpace: date axis', () => {
	it('writes the category axis as a date axis, without the category-only settings', () => {
		const minimal = reread({ kind: 'line' }, { ...MINIMAL, dateAxis: true }).plotArea.axes;
		expect(minimal.map((axis) => axis.kind)).toStrictEqual(['date', 'val']);
		const office = writeChartSpace(newChartSpace(SHEET, { kind: 'line' }, { dateAxis: true }));
		expect(office).toContain('<c:dateAx>');
		expect(office).not.toContain('<c:lblAlgn');
		expect(office).not.toContain('<c:noMultiLvlLbl');
		expect(parseChartSpace(office).issues).toStrictEqual([]);
	});

	it('leaves the value axes of a scatter chart alone', () => {
		const axes = reread({ kind: 'scatter' }, { ...MINIMAL, dateAxis: true }).plotArea.axes;
		expect(axes.map((axis) => axis.kind)).toStrictEqual(['val', 'val']);
	});
});

describe('newChartSpace: group settings', () => {
	it('starts the group with the settings its element allows', () => {
		const xml = xmlOf({ kind: 'bar' }, { ...MINIMAL, group: { gapWidth: 150, overlap: 0 } });
		expect(xml).toContain('<c:gapWidth val="150"/><c:overlap val="0"/>');
		const surface = reread({ kind: 'surface' }, { ...MINIMAL, group: { wireframe: true } });
		expect(surface.plotArea.groups[0]!.wireframe).toBe(true);
		const stock = reread({ kind: 'stock' }, { ...MINIMAL, group: { hiLowLines: true } });
		expect(stock.plotArea.groups[0]!.hiLowLines).toBe(true);
	});

	it('applies to the Office profile as well', () => {
		const xml = writeChartSpace(newChartSpace(SHEET, undefined, { group: { gapWidth: 80 } }));
		expect(xml).toContain('<c:gapWidth val="80"/>');
	});

	it('refuses a setting the element does not have, or a value out of its schema range', () => {
		const build = (target: ChartTypeTarget, group: NewChartOptions['group']) => () =>
			newChartSpace(SHEET, target, { ...MINIMAL, ...(group ? { group } : {}) });
		expect(build({ kind: 'line' }, { overlap: 10 })).toThrow(RangeError);
		expect(build({ kind: 'pie' }, { holeSize: 50 })).toThrow(RangeError);
		expect(build({ kind: 'bar' }, { gapWidth: 501 })).toThrow(RangeError);
		expect(build({ kind: 'bar' }, { gapWidth: 12.5 })).toThrow(RangeError);
		expect(build({ kind: 'doughnut' }, { holeSize: 0 })).toThrow(RangeError);
		expect(build({ kind: 'doughnut' }, { holeSize: 1 })).not.toThrow();
	});

	it('keeps a caller-owned list unshared', () => {
		const customSplit = [1, 2];
		const space = newChartSpace(SHEET, { kind: 'ofPie' }, { ...MINIMAL, group: { customSplit } });
		customSplit.push(3);
		expect(space.plotArea.groups[0]!.customSplit).toStrictEqual([1, 2]);
	});
});

describe('newChartSpace: title and legend', () => {
	it('writes no legend for null and the given position otherwise', () => {
		expect(xmlOf(undefined, { ...MINIMAL, legendPosition: null })).not.toContain('<c:legend>');
		expect(reread(undefined, { ...MINIMAL, legendPosition: 'r' }).legend).toMatchObject({
			position: 'r',
			overlay: false,
		});
		expect(reread().legend?.position).toBe('b');
	});

	it('marks the automatic title deleted without a title, and not with one', () => {
		expect(reread().autoTitleDeleted).toBe(true);
		const titled = reread(undefined, { ...MINIMAL, title: 'Sales & <costs>' });
		expect(titled.autoTitleDeleted).toBe(false);
		expect(titled.title?.text).toBe('Sales & <costs>');
		expect(xmlOf(undefined, { ...MINIMAL, title: 'T' })).toContain(
			'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>T</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>',
		);
		expect(writeChartSpace(newChartSpace(SHEET, undefined, { title: 'T' }))).toContain(
			'<a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>T</a:t></a:r>',
		);
	});
});
