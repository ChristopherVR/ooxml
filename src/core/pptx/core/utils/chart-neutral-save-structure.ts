import {
	addChartSeries,
	chartSeriesEntries,
	moveChartSeries,
	removeChartSeries,
	setChartSeriesColor,
} from '../../../chart/index';
import type { PptxChartData, PptxChartSeries } from '../types';
import { same } from './chart-neutral-save-common';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

/** Aligns the compatibility snapshot after structural operations; other families still own edits. */
export function planNeutralSeriesStructure(
	before: PptxChartData,
	after: PptxChartData,
	edits: NeutralChartEdit[],
): PptxChartData {
	if (
		before.series.length === after.series.length &&
		before.series.every((item, i) => item.idx === after.series[i]!.idx)
	)
		return before;
	const indices = before.series.map((item) => item.idx);
	if (indices.some((id) => id === undefined) || new Set(indices).size !== indices.length)
		return before;
	const retained = after.series.filter((item) => item.idx !== undefined);
	if (
		new Set(retained.map((item) => item.idx)).size !== retained.length ||
		retained.some((item) => !indices.includes(item.idx))
	)
		return before;
	// Existing series may reorder; new series append to a single-type chart without guessing groups.
	const firstNew = after.series.findIndex((item) => item.idx === undefined);
	const added = firstNew < 0 ? [] : after.series.slice(firstNew);
	if (
		added.some((item) => item.idx !== undefined) ||
		(added.length && before.chartType === 'combo')
	)
		return before;
	const basic = (series: PptxChartSeries) => ({
		name: series.name,
		values: series.values,
		...(series.color === undefined ? {} : { color: series.color }),
	});
	if (
		added.some(
			(item) =>
				!same(item, basic(item)) ||
				(item.color !== undefined && !/^#?[\da-f]{6}$/i.test(item.color)),
		)
	)
		return before;
	const wanted = retained.map((item) => item.idx!);
	const input = structuredClone(added);
	edits.push((space) => {
		let next = space;
		for (let i = before.series.length - 1; i >= 0; i--)
			if (!wanted.includes(before.series[i]!.idx!)) next = removeChartSeries(next, i);
		for (let position = 0; position < wanted.length; position++) {
			const from = chartSeriesEntries(next).findIndex(
				({ series }) => series.index === wanted[position],
			);
			if (from < 0) throw new RangeError('Chart series identity is missing');
			if (from !== position) next = moveChartSeries(next, from, position);
		}
		for (const item of input) {
			next = addChartSeries(next, {
				name: item.name,
				values: item.values,
				categories: after.categories,
			});
			if (item.color)
				next = setChartSeriesColor(next, chartSeriesEntries(next).length - 1, {
					kind: 'srgb',
					value: item.color.replace(/^#/, ''),
					transforms: [],
				});
		}
		return next;
	});
	return {
		...before,
		series: [
			...retained.map((item) => before.series.find((old) => old.idx === item.idx)!),
			...added.map(basic),
		],
	};
}
