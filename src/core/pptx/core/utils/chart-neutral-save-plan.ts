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
	setChartCategories,
	setChartLegend,
	setChartSeriesValues,
	setChartTitle,
	setChartTitleVisible,
	type ChartLegendEdit,
	type ChartLegendPosition,
	type ChartSpace,
} from '../../../chart/index';
import type { PptxChartData } from '../types';
import { same, revertStyle } from './chart-neutral-save-common';
import { planAxes } from './chart-neutral-save-axes';
import { planNeutralGroupOptions } from './chart-neutral-save-groups';
import { planNeutralSeriesStructure } from './chart-neutral-save-structure';
import { planNeutralChartType } from './chart-neutral-save-type';
import { planNeutralChartFormats } from './chart-neutral-save-format';
import { planNeutralFlatPaint } from './chart-neutral-save-flat-paint';
import { planNeutralChartText } from './chart-neutral-save-text';
import { planNeutralChartLayout } from './chart-neutral-save-layout';
import { planNeutralChartShapes } from './chart-neutral-save-shape';
import { planNeutralPointLabels } from './chart-neutral-save-point-label';
import { planNeutralChartChrome } from './chart-neutral-save-chrome';
import { planNeutralChartAnalysis } from './chart-neutral-save-analysis';
import { planNeutralLabelFormatting } from './chart-neutral-save-label-format';
import { planNeutralChartDecorations } from './chart-neutral-save-decorations';
import {
	planNeutralChartLabels,
	planNeutralChartSeriesColors,
} from './chart-neutral-save-paint-labels';

export type NeutralChartEdit = (space: ChartSpace) => ChartSpace;

export interface NeutralChartSavePlan {
	edits: NeutralChartEdit[];
	/** Structural edits need fresh workbook ranges, allocated by the package adapter. */
	rebuildWorkbook?: boolean;
	/** Whether series names, values or categories changed (the embedded workbook follows them). */
	seriesData: boolean;
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
		edits.push((space) => setChartTitle(space, title, { preserveRuns: true }));
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
	const original = before;
	before = planNeutralSeriesStructure(before, after, edits);
	planNeutralChartType(before, probe, edits);
	planTitle(before, probe, edits);
	planLegend(before, probe, edits);
	planAxes(before, probe, edits);
	planNeutralGroupOptions(before, probe, edits);
	planNeutralChartAnalysis(before, probe, edits);
	const seriesData = planSeriesData(before, probe, edits) || before !== original;
	planNeutralChartLabels(before, probe, edits);
	planNeutralLabelFormatting(before, probe, edits);
	planNeutralChartFormats(before, probe, edits);
	planNeutralChartText(before, probe, edits);
	planNeutralChartLayout(before, probe, edits);
	planNeutralChartShapes(before, probe, edits);
	planNeutralPointLabels(before, probe, edits);
	planNeutralChartChrome(before, probe, edits);
	planNeutralChartDecorations(before, probe, edits);
	planNeutralFlatPaint(before, probe, edits);
	planNeutralChartSeriesColors(before, probe, edits);
	return same(probe, before)
		? { edits, seriesData, ...(before !== original ? { rebuildWorkbook: true } : {}) }
		: undefined;
}
