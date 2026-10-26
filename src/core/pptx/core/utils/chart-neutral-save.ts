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
	chartSpaceWorkbookWrites,
	chartIssuesLostOnWrite,
	parseChartSpace,
	writeChartSpaceOverSource,
	type ChartSpace,
} from '../../../chart/index';
import type { PptxChartData } from '../types';
import { planNeutralChartSave } from './chart-neutral-save-plan';
import type { PptxChartWorkbookWrite } from './chart-xlsx-writer';

export interface NeutralChartSave {
	xml: string;
	/** Rebase these caches atomically with a new embedded worksheet before committing the part. */
	workbookModel?: ChartSpace;
	/** The embedded-workbook writes of every series, as the object-tree save collects them. */
	workbookWrites: PptxChartWorkbookWrite[];
}

/**
 * `xml` (the part as stored, which `before` describes) saved with the edits that make it `after`,
 * or `undefined` when the shared core cannot save this part or this edit yet.
 */
export function saveChartPartThroughNeutral(
	xml: string,
	before: PptxChartData,
	after: PptxChartData,
	options: { retryWorkbook?: boolean } = {},
): NeutralChartSave | undefined {
	const plan = planNeutralChartSave(before, after);
	if (!plan) return undefined;
	let space: ChartSpace;
	try {
		const parsed = parseChartSpace(xml);
		if (chartIssuesLostOnWrite(parsed.issues).length > 0) return undefined;
		space = parsed.chartSpace;
		for (const edit of plan.edits) space = edit(space);
	} catch {
		// An edit the shared operation rejects (an axis bound past the other, a missing axis) is
		// left to the object-tree writers, which report or ignore it as they always have.
		return undefined;
	}
	return {
		xml: writeChartSpaceOverSource(xml, space),
		workbookWrites: plan.seriesData || options.retryWorkbook ? chartSpaceWorkbookWrites(space) : [],
		...(plan.rebuildWorkbook ? { workbookModel: space } : {}),
	};
}
