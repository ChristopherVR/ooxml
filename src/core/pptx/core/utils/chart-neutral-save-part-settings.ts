/**
 * Chart part settings outside the plot area saved through the shared chart core
 * (docs/agnostic-core-plan.md, chart step C3): the external data auto-update flag, the colour map
 * override, the print settings, the pivot chart formats, the user shapes reference and the colour
 * style part.
 *
 * The flag and the colour mapping are shared operations over the model. The print settings and
 * pivot formats are encoded by PowerPoint's element codecs (`applyChartPrintSettings`,
 * `applyChartPivotFormats`, which also merge with the authored element) and the user shapes
 * drawing part and relationship are written by the package adapter; the plan names those as
 * `packageElements`, and the save supplies them (`saveChartPartThroughNeutral`). The colour style
 * part (`colorPalette`, `colorMethod`) is a sibling part the save loop writes before either path;
 * the chart part itself does not change for it.
 *
 * @module utils/chart-neutral-save-part-settings
 */

import { setChartColorMapOverride, setChartExternalDataAutoUpdate } from '../../../chart/index';
import type { PptxChartData } from '../types';
import { same } from './chart-neutral-save-common';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

/** Chart part elements the package adapter encodes for the neutral save. */
export type NeutralChartPackageElement = 'printSettings' | 'pivotFormats' | 'userShapes';

function revert<K extends keyof PptxChartData>(
	probe: PptxChartData,
	before: PptxChartData,
	key: K,
): void {
	if (before[key] === undefined) delete probe[key];
	else probe[key] = structuredClone(before[key]);
}

/** Plans the part settings of `probe` that differ from `before`; returns the package elements. */
export function planNeutralPartSettings(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): NeutralChartPackageElement[] {
	const elements: NeutralChartPackageElement[] = [];
	const wasAuto = before.externalData?.autoUpdate;
	const nowAuto = probe.externalData?.autoUpdate;
	if (wasAuto !== nowAuto && probe.externalData) {
		// Absent leaves the element as authored, as the object-tree writer does.
		if (nowAuto !== undefined)
			edits.push((space) => setChartExternalDataAutoUpdate(space, nowAuto));
		if (wasAuto === undefined) delete probe.externalData.autoUpdate;
		else probe.externalData.autoUpdate = wasAuto;
	}
	if (!same(before.clrMapOvr, probe.clrMapOvr) && probe.clrMapOvr !== undefined) {
		const mapping =
			probe.clrMapOvr && Object.keys(probe.clrMapOvr).length > 0 ? { ...probe.clrMapOvr } : null;
		edits.push((space) => setChartColorMapOverride(space, mapping));
		revert(probe, before, 'clrMapOvr');
	}
	if (!same(before.printSettings, probe.printSettings) && probe.printSettings !== undefined) {
		elements.push('printSettings');
		revert(probe, before, 'printSettings');
	}
	if (!same(before.pivotFormats, probe.pivotFormats) && probe.pivotFormats !== undefined) {
		elements.push('pivotFormats');
		revert(probe, before, 'pivotFormats');
	}
	if (!same(before.userShapes, probe.userShapes) && probe.userShapes !== undefined) {
		elements.push('userShapes');
		revert(probe, before, 'userShapes');
	}
	for (const key of [
		'colorPalette',
		'colorMethod',
		'colorStyleOriginalPalette',
		'colorStyleOriginalMethod',
	] as const)
		if (!same(before[key], probe[key])) revert(probe, before, key);
	return elements;
}
