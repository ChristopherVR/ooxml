import { describe, expect, it } from 'vitest';
import { columnChart } from '../../../chart/edit/test-support/column-chart';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { planNeutralChartType } from './chart-neutral-save-type';
import { same } from './chart-neutral-save-common';

const chart = (): PptxChartData =>
	({
		chartType: 'bar',
		categories: ['North', 'South', 'East'],
		series: [
			{ name: 'Plan', values: [10, 20, 30] },
			{ name: 'Actual', values: [12, 18, 33] },
		],
	}) as PptxChartData;
describe('classic per-series type save plan', () => {
	it('claims a combo override and its return to the default chart family', () => {
		const before = chart(),
			after = structuredClone(before),
			edits: NeutralChartEdit[] = [];
		after.series[1]!.seriesChartType = 'line';
		const probe = structuredClone(after);
		planNeutralChartType(before, probe, edits);
		expect(same(before, probe)).toBe(true);
		let space = columnChart();
		for (const edit of edits) space = edit(space);
		expect(space.plotArea.groups.map((group) => group.kind)).toEqual(['bar', 'line']);
		const reset: NeutralChartEdit[] = [];
		planNeutralChartType(after, structuredClone(before), reset);
		for (const edit of reset) space = edit(space);
		expect(space.plotArea.groups.map((group) => group.kind)).toEqual(['bar', 'bar']);
	});
	it('claims the SDK promotion to combo, converting only the series whose type changed', () => {
		const before = chart(),
			after = structuredClone(before),
			edits: NeutralChartEdit[] = [];
		// What `setChartSeriesChartType(element, 1, 'scatter')` leaves in the model.
		after.chartType = 'combo';
		after.series[0]!.seriesChartType = 'bar';
		after.series[1]!.seriesChartType = 'scatter';
		const probe = structuredClone(after);
		planNeutralChartType(before, probe, edits);
		expect(same(before, probe)).toBe(true);
		let space = columnChart();
		for (const edit of edits) space = edit(space);
		expect(space.plotArea.groups.map((group) => group.element)).toEqual([
			'barChart',
			'scatterChart',
		]);
		expect(space.plotArea.groups[1]!.axisIds).toEqual(space.plotArea.groups[0]!.axisIds);
	});
	it('reapplies explicit combo overrides after a whole-chart type change', () => {
		const before = chart(),
			after = structuredClone(before),
			edits: NeutralChartEdit[] = [];
		after.chartType = 'line';
		after.series[1]!.seriesChartType = 'bar';
		const probe = structuredClone(after);
		planNeutralChartType(before, probe, edits);
		expect(same(before, probe)).toBe(true);
		let space = columnChart();
		for (const edit of edits) space = edit(space);
		expect(space.plotArea.groups.map((group) => group.kind)).toEqual(['line', 'bar']);
	});
});
