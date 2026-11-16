/**
 * Chart Filters edits (`PptxChartData.filteredSeries`) saved through the shared chart core
 * (docs/agnostic-core-plan.md, chart step C3): a series hidden from the plot moves into its group's
 * `c15:filtered*Series` extension with its references and formatting (`filterChartSeries`), a
 * restored one moves back (`restoreChartFilteredSeries`) and a dropped entry leaves the part
 * (`removeChartFilteredSeries`). The object-tree writers rebuilt the extension from cached
 * literal values, losing the hidden series' formulas and formatting.
 *
 * Runs before the series structure planner: it returns the chart as it stands after these moves,
 * so the structure planner sees no series added or removed for them.
 *
 * @module utils/chart-neutral-save-filtered
 */

import {
	filterChartSeries,
	removeChartFilteredSeries,
	restoreChartFilteredSeries,
} from '../../../chart/index';
import type { PptxChartData, PptxChartFilteredSeries } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

/** The parsed fields of an entry (`originalSeries` is an editing-session snapshot, never saved). */
const entryKey = (entry: PptxChartFilteredSeries) => {
	const saved: Partial<PptxChartFilteredSeries> = { ...entry };
	delete saved.originalSeries;
	return JSON.stringify(saved);
};

/** `before` after the filter moves `after` asks for, or `before` itself when there are none. */
export function planNeutralFilteredSeries(
	before: PptxChartData,
	after: PptxChartData,
	edits: NeutralChartEdit[],
): PptxChartData {
	const was = before.filteredSeries ?? [];
	const now = after.filteredSeries ?? [];
	if (was.length === 0 && now.length === 0) return before;
	const wasByIdx = new Map(was.map((entry) => [entry.idx, entry]));
	const nowIdx = new Set(now.map((entry) => entry.idx));
	const visibleBefore = new Set(before.series.map((series) => series.idx));
	const visibleAfter = new Set(after.series.map((series) => series.idx));
	const hidden: number[] = [];
	for (const entry of now) {
		const old = wasByIdx.get(entry.idx);
		if (old) {
			// A kept entry must be unchanged; the object-tree path decides any other edit.
			if (entryKey(old) !== entryKey(entry)) return before;
			continue;
		}
		if (!visibleBefore.has(entry.idx) || visibleAfter.has(entry.idx)) return before;
		hidden.push(entry.idx);
	}
	const restored: number[] = [];
	const removed: number[] = [];
	for (const entry of was) {
		if (nowIdx.has(entry.idx)) continue;
		if (visibleAfter.has(entry.idx)) restored.push(entry.idx);
		else removed.push(entry.idx);
	}
	if (hidden.length + restored.length + removed.length === 0) return before;
	if (before.series.some((series) => series.idx === undefined)) return before;
	// The series the moves leave alone keep their order; a reorder is the structure planner's.
	const keptOrder = after.series
		.map((item) => item.idx)
		.filter((idx) => visibleBefore.has(idx) && !hidden.includes(idx!));
	const beforeOrder = before.series.map((item) => item.idx).filter((idx) => !hidden.includes(idx!));
	if (JSON.stringify(keptOrder) !== JSON.stringify(beforeOrder)) return before;
	edits.push((space) => {
		let next = space;
		// Restore first when the edit swaps the only plotted series with a filtered one.
		for (const idx of restored) next = restoreChartFilteredSeries(next, idx);
		for (const idx of hidden) next = filterChartSeries(next, idx);
		for (const idx of removed) next = removeChartFilteredSeries(next, idx);
		return next;
	});
	// The chart as it stands after the moves: hidden series gone, restored ones in the position
	// the edited model gives them (their data comes back from the part, as written).
	const series = after.series.filter(
		(item) =>
			restored.includes(item.idx!) || (visibleBefore.has(item.idx) && !hidden.includes(item.idx!)),
	);
	const kept = new Map(before.series.map((item) => [item.idx, item]));
	return {
		...before,
		series: series.map((item) => (restored.includes(item.idx!) ? item : kept.get(item.idx)!)),
		...(after.filteredSeries === undefined
			? { filteredSeries: undefined }
			: { filteredSeries: after.filteredSeries }),
	};
}
