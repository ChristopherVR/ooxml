import { describe, expect, it } from 'vitest';
import type { DrawingColor } from '../../drawingml/types';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { setChartAxis } from './axes';
import { setChartDataLabels } from './legend-labels';
import { setChartGroupOptions } from './group-options';
import { applyChartEditIntent } from './intent';
import { chartFormatTargetKey, parseChartFormatTargetKey } from './target';
import { chartTargetShape, setChartTargetFill, setChartTargetLine } from './target-paint';
import { columnChart } from './test-support/column-chart';

const red: DrawingColor = { kind: 'srgb', value: 'FF0000', transforms: [] };
const reload = (space: ReturnType<typeof columnChart>) => {
	const parsed = parseChartSpace(writeChartSpace(space));
	expect(parsed.issues).toEqual([]);
	return parsed.chartSpace;
};

describe('format target keys', () => {
	it('round-trips every target kind and refuses malformed keys', () => {
		for (const key of ['chartArea', 'legend', 'axis:1', 'minorGridlines:0', 'series:2'])
			expect(chartFormatTargetKey(parseChartFormatTargetKey(key)!)).toBe(key);
		expect(parseChartFormatTargetKey('dataLabels')).toEqual({ kind: 'dataLabels' });
		expect(parseChartFormatTargetKey('dataLabels:3')).toEqual({
			kind: 'dataLabels',
			seriesIndex: 3,
		});
		for (const key of ['axis', 'series:-1', 'series:1.5', 'legend:0', 'nope'])
			expect(parseChartFormatTargetKey(key)).toBeUndefined();
	});
});

describe('target paint', () => {
	it('fills the plot area, a series and the data labels through one call', () => {
		let space = setChartTargetFill(
			columnChart(),
			{ kind: 'plotArea' },
			{ kind: 'solid', color: red },
		);
		space = setChartTargetFill(space, { kind: 'series', seriesIndex: 1 }, { kind: 'none' });
		space = setChartDataLabels(space, { show: true });
		space = setChartTargetFill(space, { kind: 'dataLabels' }, { kind: 'solid', color: red });
		const back = reload(space);
		expect(chartTargetShape(back, { kind: 'plotArea' })?.fill).toEqual({
			kind: 'solid',
			color: red,
		});
		expect(chartTargetShape(back, { kind: 'series', seriesIndex: 1 })?.fill).toEqual({
			kind: 'none',
		});
		expect(chartTargetShape(back, { kind: 'dataLabels' })?.fill).toEqual({
			kind: 'solid',
			color: red,
		});
	});

	it('refuses a fill on a line target and labels that are not shown', () => {
		expect(() =>
			setChartTargetFill(columnChart(), { kind: 'axis', axisIndex: 0 }, undefined),
		).toThrow(/no fill/);
		expect(() =>
			setChartTargetFill(columnChart(), { kind: 'dataLabels', seriesIndex: 0 }, undefined),
		).toThrow(/no data labels/);
	});

	it('styles an axis line and turns gridlines on when they are styled', () => {
		const line = { widthEmu: 12700, fill: { kind: 'solid' as const, color: red } };
		let space = setChartTargetLine(columnChart(), { kind: 'axis', axisIndex: 0 }, line);
		space = setChartTargetLine(space, { kind: 'minorGridlines', axisIndex: 1 }, line);
		space = setChartTargetLine(space, { kind: 'axis', axisIndex: 0 }, { dash: 'dash' }, true);
		const back = reload(space);
		expect(back.plotArea.axes[0]!.spPr?.line).toMatchObject({ widthEmu: 12700, dash: 'dash' });
		expect(back.plotArea.axes[1]).toMatchObject({ minorGridlines: true });
		expect(back.plotArea.axes[1]!.minorGridlinesSpPr?.line?.widthEmu).toBe(12700);
	});
});

describe('group options', () => {
	it('sets gap width and overlap on the group of a series', () => {
		const back = reload(setChartGroupOptions(columnChart(), 1, { gapWidth: 40, overlap: 25 }));
		expect(back.plotArea.groups[0]).toMatchObject({ gapWidth: 40, overlap: 25 });
	});

	it('refuses values out of range and settings the group does not have', () => {
		expect(() => setChartGroupOptions(columnChart(), 0, { gapWidth: 501 })).toThrow(/0 to 500/);
		expect(() => setChartGroupOptions(columnChart(), 0, { holeSize: 50 })).toThrow(/no holeSize/);
		expect(() => setChartGroupOptions(columnChart(), 5, { gapWidth: 10 })).toThrow(
			/No chart series/,
		);
	});
});

describe('applyChartEditIntent', () => {
	it('runs each intent as the operation it names', () => {
		const start = columnChart();
		expect(
			applyChartEditIntent(start, { op: 'axis', axisIndex: 1, edit: { max: 50 } }).plotArea.axes[1]!
				.scaling.max,
		).toBe(50);
		expect(applyChartEditIntent(start, { op: 'axis', axisIndex: 1, edit: { max: 50 } })).toEqual(
			setChartAxis(start, 'val', { max: 50 }),
		);
		expect(
			applyChartEditIntent(start, { op: 'legend', edit: { position: 't' } }).legend?.position,
		).toBe('t');
		const labelled = applyChartEditIntent(start, {
			op: 'dataLabels',
			seriesIndex: 0,
			edit: { showValue: true },
		});
		expect(labelled.plotArea.groups[0]!.series[0]!.dataLabels?.showValue).toBe(true);
		expect(
			applyChartEditIntent(start, { op: 'groupOptions', seriesIndex: 0, edit: { gapWidth: 80 } })
				.plotArea.groups[0]!.gapWidth,
		).toBe(80);
		const solid = applyChartEditIntent(start, {
			op: 'fill',
			target: { kind: 'chartArea' },
			fill: { kind: 'solid', color: red },
		});
		const clear = applyChartEditIntent(solid, {
			op: 'transparency',
			target: { kind: 'chartArea' },
			paint: 'fill',
			percent: 25,
		});
		expect(clear.spPr?.fill).toEqual({
			kind: 'solid',
			color: { ...red, transforms: [{ name: 'alpha', value: '75000' }] },
		});
		expect(start.spPr).toBeUndefined();
	});
});
