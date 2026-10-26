// Claims typed categories and independent XY/bubble vectors before the flat data planner runs.
import {
	setChartTypedCategories,
	setChartNumericCategories,
	setChartCategoryLevels,
	setChartSeriesXValues,
	setChartSeriesBubbleSizes,
} from '../../../chart';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';
/** Run before planSeriesData; use changed for seriesData and rebuildWorkbook for fresh ranges. */
export function planNeutralDataVectors(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): { changed: boolean; rebuildWorkbook: boolean } {
	let changed = false;
	let rebuildWorkbook = false;
	const claim = (rebuild = true) => {
		changed = true;
		rebuildWorkbook ||= rebuild;
	};
	if (
		!same(before.dateCategories, probe.dateCategories) &&
		probe.dateCategories &&
		Boolean(before.dateCategories?.date1904) === Boolean(probe.dateCategories.date1904)
	) {
		const { values } = probe.dateCategories;
		const formatCode =
			probe.dateCategories.formatCode ??
			(before.dateCategories?.formatCode ? 'General' : undefined);
		edits.push((space) => setChartNumericCategories(space, values, formatCode));
		probe.dateCategories = before.dateCategories;
		probe.categories = before.categories;
		claim();
	}
	if (!same(before.categoryLevels, probe.categoryLevels) && probe.categoryLevels?.length) {
		const levels = probe.categoryLevels;
		edits.push((space) => setChartCategoryLevels(space, levels));
		probe.categoryLevels = before.categoryLevels;
		probe.categories = before.categories;
		claim();
	} else if (!same(before.categories, probe.categories) && !probe.dateCategories) {
		const values = [...probe.categories];
		const rebuild =
			Boolean(before.categoryLevels) ||
			values.length !== before.categories.length ||
			['scatter', 'bubble'].includes(before.chartType);
		edits.push((space) => setChartTypedCategories(space, values));
		probe.categories = before.categories;
		claim(rebuild);
	}
	for (const [index, series] of probe.series.entries()) {
		const old = before.series[index];
		if (!old) continue;
		for (const [key, operation] of [
			['xValues', setChartSeriesXValues],
			['bubbleSizes', setChartSeriesBubbleSizes],
		] as const) {
			if (same(old[key], series[key]) || !series[key]) continue;
			const values = [...series[key]];
			edits.push((space) => operation(space, index, values));
			if (old[key] === undefined) delete series[key];
			else series[key] = old[key];
			claim();
		}
	}
	return { changed, rebuildWorkbook };
}
