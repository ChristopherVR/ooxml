/**
 * Saves an edited classic chart part through the shared chart core
 * (docs/agnostic-core-plan.md, chart step C3): the part as stored is read by
 * `parseChartSpace`, the edits planned by `planNeutralChartSave` are applied
 * with the `chart/edit` operations and the result is written by
 * the shared source-backed writer, retaining no-op bytes and carried top-level children.
 * A part the neutral parser reports content it cannot preserve in, or an edit
 * no family claims yet, returns `undefined` so the caller keeps the
 * object-tree writers for it.
 *
 * @module utils/chart-neutral-save
 */

import {
	chartIssuesLostOnWrite,
	setChartPivotFormatsXml,
	setChartPrintSettingsXml,
	setChartUserShapesReference,
	type ChartSpace,
} from '../../../chart/index';
import {
	readChartSpaceSource,
	writeParsedChartSpace,
	type ChartSpaceSource,
} from '../../../chart/keep-source';
import type { PptxChartData } from '../types';
import { planNeutralChartSave } from './chart-neutral-save-plan';

/**
 * Chart part elements the package adapter encoded for a save: the print settings and pivot formats
 * (`null` removes the element) and the user shapes relationship id (`null` removes the reference).
 */
export interface NeutralChartPackageElements {
	printSettings?: string | null;
	pivotFormats?: string | null;
	/** `undefined` keeps the authored reference (the drawing on disk already matched). */
	userShapes?: string | null | undefined;
}

export interface NeutralChartSave {
	xml: string;
	/** Rebase these caches atomically with a new embedded worksheet before committing the part. */
	workbookModel?: ChartSpace;
	/** The workbook already holds the chart's data, on a sheet made for this save. */
	workbookRebased?: boolean;
}

/**
 * `xml` (the part as stored, which `before` describes) saved with the edits that make it `after`,
 * or `undefined` when the shared core cannot save this part or this edit yet.
 */
export function saveChartPartThroughNeutral(
	xml: string,
	before: PptxChartData,
	after: PptxChartData,
	options: {
		elements?: NeutralChartPackageElements | undefined;
		readSource?: ((xml: string) => ChartSpaceSource) | undefined;
	} = {},
): NeutralChartSave | undefined {
	const plan = planNeutralChartSave(before, after);
	if (!plan) return undefined;
	const elements = options.elements ?? {};
	// An element the plan needs and the package adapter did not encode leaves the object-tree path.
	if (plan.packageElements?.some((name) => !Object.hasOwn(elements, name))) return undefined;
	let space: ChartSpace;
	let written: string;
	try {
		const parsed = (options.readSource ?? readChartSpaceSource)(xml);
		if (chartIssuesLostOnWrite(parsed.issues).length > 0) return undefined;
		space = parsed.chartSpace;
		for (const edit of plan.edits) space = edit(space);
		for (const name of plan.packageElements ?? []) {
			if (name === 'printSettings')
				space = setChartPrintSettingsXml(space, elements.printSettings!);
			if (name === 'pivotFormats') space = setChartPivotFormatsXml(space, elements.pivotFormats!);
			if (name === 'userShapes' && elements.userShapes !== undefined)
				space = setChartUserShapesReference(space, elements.userShapes);
		}
		written = writeParsedChartSpace(parsed, space);
	} catch {
		// An edit the shared operation rejects (an axis bound past the other, a missing axis), or a
		// model the shared writer refuses to write, is left to the object-tree writers, which report
		// or ignore it as they always have, instead of losing the whole part's save.
		return undefined;
	}
	return {
		xml: written,
		...(plan.rebuildWorkbook ? { workbookModel: space } : {}),
	};
}
