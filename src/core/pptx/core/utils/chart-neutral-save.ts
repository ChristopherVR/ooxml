/**
 * Saves an edited classic chart part through the shared chart core
 * (docs/agnostic-core-plan.md, chart step C3): the part as stored is read by
 * `parseChartSpace`, the edits planned by `planNeutralChartSave` are applied
 * with the `chart/edit` operations and the result is written by
 * `writeChartSpace`, which round-trips everything the model holds byte for
 * byte. A part the neutral parser reports anything unmodelled in, or an edit
 * no family claims yet, returns `undefined` so the caller keeps the
 * object-tree writers for it.
 *
 * @module utils/chart-neutral-save
 */

import {
	chartSeriesEntries,
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
	/** The embedded-workbook writes of every series, as the object-tree save collects them. */
	workbookWrites: PptxChartWorkbookWrite[];
}

/**
 * The workbook writes that keep "Edit Data in Excel" in step with the saved caches: each
 * referenced series name, text category range and value range gets the model's values (the same
 * writes, and the same value text, as the object-tree save loop pushes).
 */
function workbookWrites(space: ChartSpace, chart: PptxChartData): PptxChartWorkbookWrite[] {
	const writes: PptxChartWorkbookWrite[] = [];
	const dateValues = chart.dateCategories?.values.map(String);
	chartSeriesEntries(space).forEach(({ series }, i) => {
		const data = chart.series[i];
		if (!data) return;
		const name = series.tx?.reference;
		if (name?.kind === 'strRef' && name.formula)
			writes.push({ formula: name.formula, isNumeric: false, values: [data.name] });
		const categories = series.categories ?? series.xValues;
		const categoryKind = dateValues ? 'numRef' : 'strRef';
		if (categories?.kind === categoryKind && categories.formula)
			writes.push({
				formula: categories.formula,
				isNumeric: Boolean(dateValues),
				values: dateValues ?? chart.categories,
			});
		const values = series.values ?? series.yValues;
		if (values?.kind === 'numRef' && values.formula)
			writes.push({
				formula: values.formula,
				isNumeric: true,
				values: data.values.map((value, index) => (data.blanks?.[index] ? '' : String(value))),
			});
	});
	return writes;
}

/**
 * `xml` (the part as stored, which `before` describes) saved with the edits that make it `after`,
 * or `undefined` when the shared core cannot save this part or this edit yet.
 */
export function saveChartPartThroughNeutral(
	xml: string,
	before: PptxChartData,
	after: PptxChartData,
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
		workbookWrites: plan.seriesData ? workbookWrites(space, after) : [],
	};
}
