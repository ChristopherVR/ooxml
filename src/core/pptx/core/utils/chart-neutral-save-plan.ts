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
	setChartLegend,
	setChartTitle,
	setChartTitleVisible,
	type ChartLegendEdit,
	type ChartLegendPosition,
	type ChartSpace,
} from '../../../chart/index';
import type { PptxChartData, PptxChartStyle } from '../types';

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
	return same(probe, before) ? { edits, seriesData: false } : undefined;
}
