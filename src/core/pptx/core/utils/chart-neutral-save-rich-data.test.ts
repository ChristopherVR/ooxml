import { expect, it } from 'vitest';
import { newChartSpace, chartSpaceSeries, rebaseChartData } from '../../../chart';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { planNeutralTitleRuns } from './chart-neutral-save-title-runs';
import { planNeutralDataVectors } from './chart-neutral-save-data-vectors';
const before: PptxChartData = {
	chartType: 'bar',
	categories: ['A', 'B'],
	series: [{ name: 'Sales', values: [1, 2] }],
	title: 'First',
	titleRuns: [
		{ text: 'First', bold: true },
		{ text: 'Second', italic: true },
	],
};
it('claims an explicit rich-run edit, including inherited boolean resets', () => {
	const probe = structuredClone(before);
	probe.titleRuns![1] = { text: 'Tail', fontSize: 24, color: '#123456' };
	const edits: NeutralChartEdit[] = [];
	planNeutralTitleRuns(before, probe, edits);
	expect(probe).toEqual(before);
	expect(edits).toHaveLength(1);
	const space = edits[0]!(
		newChartSpace(
			{ categories: ['A', 'B'], series: before.series },
			{ kind: 'bar' },
			{ title: 'FirstSecond' },
		),
	);
	expect(space.title!.text).toBe('FirstTail');
	expect(space.title!.tx!.rich!.sourceXml).toContain('sz="2400"');
});
it('claims date cache edits and requests a fresh workbook range without changing the date system', () => {
	const old = {
		...structuredClone(before),
		dateCategories: { values: [45000, 45001], date1904: false, formatCode: 'yyyy-mm-dd' },
	};
	const probe = structuredClone(old);
	probe.dateCategories.values = [45002, 45003];
	probe.categories = ['New 1', 'New 2'];
	const edits: NeutralChartEdit[] = [];
	expect(planNeutralDataVectors(old, probe, edits)).toEqual({
		changed: true,
		rebuildWorkbook: true,
	});
	expect(probe).toEqual(old);
	let space = newChartSpace({ categories: ['A', 'B'], series: before.series }, { kind: 'bar' });
	for (const edit of edits) space = edit(space);
	expect(chartSpaceSeries(space)[0]!.categories!.cache!.points.map((p) => p.value)).toEqual([
		'45002',
		'45003',
	]);
	expect(rebaseChartData(space, 'SafeData').rows[1]![0]).toBe(45002);
});
it('claims per-series X and bubble sizes independently and refuses an implicit date-system switch', () => {
	const old = {
		...structuredClone(before),
		chartType: 'bubble' as const,
		series: [{ name: 'Sales', values: [1, 2], xValues: [3, 4], bubbleSizes: [5, 6] }],
	};
	const probe = structuredClone(old);
	probe.series[0]!.xValues = [7, 8];
	probe.series[0]!.bubbleSizes = [9, 10];
	const edits: NeutralChartEdit[] = [];
	expect(planNeutralDataVectors(old, probe, edits).rebuildWorkbook).toBe(true);
	expect(probe).toEqual(old);
	let space = newChartSpace({ categories: ['3', '4'], series: before.series }, { kind: 'bubble' });
	for (const edit of edits) space = edit(space);
	expect(rebaseChartData(space, 'SafeData').rows[1]).toEqual([7, 1, 9]);
	const date = { ...structuredClone(before), dateCategories: { values: [1, 2], date1904: false } };
	const changed = structuredClone(date);
	changed.dateCategories.date1904 = true;
	expect(planNeutralDataVectors(date, changed, []).changed).toBe(false);
	expect(changed.dateCategories.date1904).toBe(true);
});
