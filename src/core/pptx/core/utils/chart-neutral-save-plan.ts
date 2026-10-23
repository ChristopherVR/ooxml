/**
 * Plans how an edited PowerPoint chart is saved through the shared chart core
 * (docs/agnostic-core-plan.md, chart step C3): the difference between the
 * chart as loaded (`before`) and as edited (`after`) is turned into
 * `chart/edit` operations on the part's `ChartSpace`, one feature family at a
 * time. A family claims its fields only when the shared operation writes
 * what the PowerPoint writer wrote for that edit (proven by
 * `__tests__/integration/chart-save-parity.test.ts`); when any changed field
 * is left unclaimed the plan is `undefined` and the part is saved by the
 * object-tree writers as before.
 *
 * @module utils/chart-neutral-save-plan
 */

import {
	renameChartSeries,
	setChartAxis,
	setChartCategories,
	setChartLegend,
	setChartSeriesValues,
	setChartTitle,
	setChartTitleVisible,
	type ChartAxisEdit,
	type ChartLegendEdit,
	type ChartLegendPosition,
	type ChartSpace,
} from '../../../chart/index';
import type { PptxChartAxisFormatting, PptxChartData, PptxChartStyle } from '../types';

export type NeutralChartEdit = (space: ChartSpace) => ChartSpace;

export interface NeutralChartSavePlan {
	edits: NeutralChartEdit[];
	/** Whether series names, values or categories changed (the embedded workbook follows them). */
	seriesData: boolean;
}

/** JSON of a value with sorted keys, `undefined` members and empty objects left out. */
function canonical(value: unknown): unknown {
	if (Array.isArray(value)) return value.map((item) => canonical(item) ?? null);
	if (value && typeof value === 'object') {
		const entries = Object.keys(value)
			.sort()
			.flatMap((key) => {
				const item = canonical((value as Record<string, unknown>)[key]);
				return item === undefined ? [] : [[key, item] as const];
			});
		return entries.length > 0 ? Object.fromEntries(entries) : undefined;
	}
	return value;
}

const same = (a: unknown, b: unknown): boolean =>
	JSON.stringify(canonical(a)) === JSON.stringify(canonical(b));

type StyleKey = keyof PptxChartStyle;

/** Copies `keys` of `before.style` onto `probe.style` (the fields a family has claimed). */
function revertStyle(probe: PptxChartData, before: PptxChartData, keys: readonly StyleKey[]): void {
	const style: Record<string, unknown> = { ...probe.style };
	for (const key of keys) {
		const value = before.style?.[key];
		if (value === undefined) delete style[key];
		else style[key] = value;
	}
	probe.style = style as PptxChartStyle;
}

/** Title text and visibility (`applyChartTitleToXml`'s flat-text cases). */
function planTitle(before: PptxChartData, probe: PptxChartData, edits: NeutralChartEdit[]): void {
	const hadTitle = before.style?.hasTitle;
	const hasTitle = probe.style?.hasTitle;
	if (before.title === probe.title && hadTitle === hasTitle) return;
	const title = probe.title;
	if (hasTitle === false || (title === '' && hasTitle !== true)) {
		edits.push((space) => setChartTitle(space, undefined));
	} else if (title !== undefined && title !== before.title) {
		// A multi-run title is realigned run by run by the PowerPoint writer; not claimed yet.
		if ((before.titleRuns?.length ?? 0) > 1) return;
		edits.push((space) => setChartTitle(space, title));
	} else if (hasTitle === true) {
		edits.push((space) => setChartTitleVisible(space, true));
	}
	probe.title = before.title;
	revertStyle(probe, before, ['hasTitle']);
}

const LEGEND_POSITIONS = new Set<string>(['b', 'tr', 'l', 'r', 't']);

/** Legend visibility, position and overlay (`applyChartLegendToXml` without entries). */
function planLegend(before: PptxChartData, probe: PptxChartData, edits: NeutralChartEdit[]): void {
	const was = before.style ?? {};
	const now = probe.style ?? {};
	const keys = ['hasLegend', 'legendPosition', 'legendOverlay'] as const;
	if (keys.every((key) => was[key] === now[key])) return;
	const edit: ChartLegendEdit = {};
	if (now.hasLegend === false) {
		edit.visible = false;
	} else if (now.hasLegend === true) {
		const shown = was.hasLegend !== true;
		if (shown) edit.visible = true;
		if (now.legendPosition !== undefined && (shown || now.legendPosition !== was.legendPosition)) {
			if (!LEGEND_POSITIONS.has(now.legendPosition)) return;
			edit.position = now.legendPosition as ChartLegendPosition;
		}
		if (now.legendOverlay !== undefined && now.legendOverlay !== was.legendOverlay)
			edit.overlay = now.legendOverlay;
	}
	// Without `hasLegend` the PowerPoint writer leaves the legend alone, and so does this plan.
	if (Object.keys(edit).length > 0) edits.push((space) => setChartLegend(space, edit));
	revertStyle(probe, before, keys);
}

/** The axis fields the shared `setChartAxis` writes as the PowerPoint axis writers do. */
function axisEdit(
	was: PptxChartAxisFormatting,
	now: PptxChartAxisFormatting,
): { edit: ChartAxisEdit; claimed: (keyof PptxChartAxisFormatting)[] } {
	const edit: ChartAxisEdit = {};
	const claimed: (keyof PptxChartAxisFormatting)[] = [];
	for (const key of ['min', 'max', 'majorUnit', 'minorUnit'] as const) {
		if (was[key] === now[key]) continue;
		edit[key] = now[key] ?? null;
		claimed.push(key);
	}
	if (Boolean(was.deleted) !== Boolean(now.deleted)) {
		edit.visible = !now.deleted;
		claimed.push('deleted');
	}
	if (!same(was.numFmt, now.numFmt) && now.numFmt && !now.numFmt.sourceLinked) {
		edit.numberFormat = now.numFmt.formatCode;
		claimed.push('numFmt');
	}
	if (was.orientation !== now.orientation && now.orientation !== undefined) {
		edit.reverse = now.orientation === 'maxMin';
		claimed.push('orientation');
	}
	if (was.tickLblPos !== now.tickLblPos && now.tickLblPos !== undefined) {
		edit.tickLabelPosition = now.tickLblPos;
		claimed.push('tickLblPos');
	}
	return { edit, claimed };
}

/** Value and category axis scaling, visibility, number format, direction and labels. */
function planAxes(before: PptxChartData, probe: PptxChartData, edits: NeutralChartEdit[]): void {
	const was = before.axes ?? [];
	const now = probe.axes ?? [];
	if (was.length !== now.length) return;
	now.forEach((axis, i) => {
		const previous = was[i]!;
		const id = previous.axisId;
		if (id === undefined || axis.axisId !== id || axis.axisType !== previous.axisType) return;
		const { edit, claimed } = axisEdit(previous, axis);
		if (claimed.length === 0) return;
		edits.push((space) => setChartAxis(space, { id }, edit));
		const reverted: Record<string, unknown> = { ...axis };
		for (const key of claimed) {
			if (previous[key] === undefined) delete reverted[key];
			else reverted[key] = previous[key];
		}
		now[i] = reverted as unknown as PptxChartAxisFormatting;
	});
}

/** Series names, values and text categories (the cache writes of the save loop). */
function planSeriesData(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): boolean {
	if (before.series.length !== probe.series.length) return false;
	let changed = false;
	probe.series.forEach((series, i) => {
		const previous = before.series[i]!;
		if (series.name !== previous.name) {
			const name = series.name;
			edits.push((space) => renameChartSeries(space, i, name, { keepReference: true }));
			series.name = previous.name;
			changed = true;
		}
		if (!same(series.values, previous.values) || !same(series.blanks, previous.blanks)) {
			const values = series.values.map((value, point) => (series.blanks?.[point] ? null : value));
			edits.push((space) => setChartSeriesValues(space, i, values));
			series.values = previous.values;
			if (previous.blanks === undefined) delete series.blanks;
			else series.blanks = previous.blanks;
			changed = true;
		}
	});
	const textCategories =
		!before.categoryLevels &&
		!before.dateCategories &&
		!['scatter', 'bubble'].includes(before.chartType) &&
		!before.series.some((s) => s.seriesChartType === 'scatter' || s.seriesChartType === 'bubble');
	if (!same(before.categories, probe.categories) && textCategories) {
		const categories = [...probe.categories];
		edits.push((space) => setChartCategories(space, categories));
		probe.categories = before.categories;
		changed = true;
	}
	return changed;
}

/**
 * The shared-core edits that turn the part saved as `before` into `after`, or
 * `undefined` when `after` changes something no family claims yet.
 */
export function planNeutralChartSave(
	before: PptxChartData,
	after: PptxChartData,
): NeutralChartSavePlan | undefined {
	const probe = structuredClone(after);
	const edits: NeutralChartEdit[] = [];
	planTitle(before, probe, edits);
	planLegend(before, probe, edits);
	planAxes(before, probe, edits);
	const seriesData = planSeriesData(before, probe, edits);
	return same(probe, before) ? { edits, seriesData } : undefined;
}
