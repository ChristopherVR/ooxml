/**
 * Chart group subtype settings saved through the shared chart core (docs/agnostic-core-plan.md,
 * chart step C3): the 3-D bar shape, the 3-D gap depth, the radar style, the surface wireframe flag,
 * the pie-of-pie split options and the bubble options, mapped onto `setChartGroupSubtype`.
 *
 * Each field is claimed only under the condition on which the object-tree writers write it (the
 * chart type of the edited model, see `PptxHandlerRuntimeSaveDataSerialization.ts`) and with the
 * same removal rule: a field those writers remove when it is cleared is removed here, one they
 * leave in place when cleared is left unclaimed so the object-tree path keeps deciding it.
 *
 * @module utils/chart-neutral-save-subtype
 */

import {
	chartSeriesEntries,
	setChartGroupOptions,
	setChartGroupSubtype,
	type ChartGroupSubtypeEdit,
	type ChartSpace,
} from '../../../chart/index';
import type { PptxChartData } from '../types';
import { same } from './chart-neutral-save-common';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

/** Applies `apply` (to a series of the group) once for every group whose element is in `elements`. */
function onGroups(
	elements: readonly string[],
	apply: (space: ChartSpace, seriesIndex: number) => ChartSpace,
): NeutralChartEdit {
	return (space: ChartSpace) => {
		const visited = new Set<number>();
		let next = space;
		for (const [index, { group, groupIndex }] of chartSeriesEntries(space).entries()) {
			if (visited.has(groupIndex) || !elements.includes(group.element)) continue;
			visited.add(groupIndex);
			next = apply(next, index);
		}
		return next;
	};
}

/** Copies `key` of `before` onto `probe` (the field is claimed). */
function revert<K extends keyof PptxChartData>(
	probe: PptxChartData,
	before: PptxChartData,
	key: K,
): void {
	if (before[key] === undefined) delete probe[key];
	else probe[key] = structuredClone(before[key]);
}

/** Fields the object-tree writers set from the model, removing the element when cleared. */
const REMOVABLE = [
	{ key: 'barShape', field: 'shape', types: ['bar3D'], elements: ['bar3DChart'] },
	{
		key: 'gapDepth',
		field: 'gapDepth',
		types: ['bar3D', 'area3D', 'line3D'],
		elements: ['bar3DChart', 'area3DChart', 'line3DChart'],
	},
	{ key: 'radarStyle', field: 'radarStyle', types: ['radar'], elements: ['radarChart'] },
	{
		key: 'wireframe',
		field: 'wireframe',
		types: ['surface'],
		elements: ['surfaceChart', 'surface3DChart'],
	},
] as const;

const subtype = (edit: ChartGroupSubtypeEdit) => (space: ChartSpace, index: number) =>
	setChartGroupSubtype(space, index, edit);

function planOfPie(before: PptxChartData, probe: PptxChartData, edits: NeutralChartEdit[]): void {
	const was = before.ofPieOptions;
	const now = probe.ofPieOptions;
	if (same(was, now)) return;
	// Cleared options are left to the object-tree path (it keeps the authored elements).
	if (!now || probe.chartType !== 'ofPie' || before.chartType !== 'ofPie') return;
	const edit: ChartGroupSubtypeEdit = {};
	if (now.ofPieType !== was?.ofPieType) edit.ofPieType = now.ofPieType;
	if (now.splitType !== was?.splitType) edit.splitType = now.splitType ?? null;
	if (now.splitPos !== was?.splitPos) edit.splitPosition = now.splitPos ?? null;
	if (now.secondPieSize !== was?.secondPieSize) edit.secondPieSize = now.secondPieSize ?? null;
	if (now.serLines !== was?.serLines) {
		// `serLines` absent keeps the authored element, as the object-tree writer does.
		if (now.serLines === undefined) return;
		edit.seriesLines = now.serLines ? true : null;
	}
	if (!same(now.custSplit, was?.custSplit)) {
		// An empty or absent list keeps the authored split points, as the object-tree writer does.
		if (!now.custSplit || now.custSplit.length === 0) return;
		edit.customSplit = [...now.custSplit];
	}
	const gapWidth = now.gapWidth !== was?.gapWidth ? (now.gapWidth ?? null) : undefined;
	// `c:gapWidth` of a pie of pie is the gap between its two plots (a ranged group option).
	edits.push(
		onGroups(['ofPieChart'], (space, index) => {
			const next = setChartGroupSubtype(space, index, edit);
			return gapWidth === undefined ? next : setChartGroupOptions(next, index, { gapWidth });
		}),
	);
	revert(probe, before, 'ofPieOptions');
}

function planBubble(before: PptxChartData, probe: PptxChartData, edits: NeutralChartEdit[]): void {
	const was = before.bubbleOptions ?? {};
	const now = probe.bubbleOptions ?? {};
	if (same(was, now)) return;
	if (probe.chartType !== 'bubble' || before.chartType !== 'bubble') return;
	const edit: ChartGroupSubtypeEdit = {};
	for (const key of ['bubble3D', 'bubbleScale', 'showNegativeBubbles', 'sizeRepresents'] as const) {
		if (now[key] === was[key]) continue;
		// The object-tree writer sets defined options only; a cleared one keeps its element.
		if (now[key] === undefined) return;
		(edit as Record<string, unknown>)[key] = now[key];
	}
	edits.push(onGroups(['bubbleChart'], subtype(edit)));
	revert(probe, before, 'bubbleOptions');
}

/** Plans the subtype settings of `probe` that differ from `before`. */
export function planNeutralGroupSubtype(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	for (const { key, field, types, elements } of REMOVABLE) {
		if (same(before[key], probe[key])) continue;
		if (!(types as readonly string[]).includes(probe.chartType)) continue;
		// The type planner owns a part whose groups change element; this family edits kept groups.
		if (before.chartType !== probe.chartType) continue;
		edits.push(onGroups(elements, subtype({ [field]: probe[key] ?? null })));
		revert(probe, before, key);
	}
	planOfPie(before, probe, edits);
	planBubble(before, probe, edits);
}
