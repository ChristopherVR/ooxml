/**
 * A new classic chart asking for an embedded workbook (`PptxChartData.embedWorkbook`): the part
 * the shared chart core writes (`writeNewChartPart`) with its literal data moved into the
 * workbook (`newChartWorkbookLayout`), so names, categories and values become cell references
 * that keep their caches, and the worksheet rows the workbook holds. The legacy generator writes
 * literals only, so a chart the shared core cannot write is refused with the reason rather than
 * saved without its workbook.
 *
 * @module utils/chart-new-part-workbook
 */
import {
	DEFAULT_CHART_SHEET_NAME,
	chartIssuesLostOnWrite,
	newChartWorkbookLayout,
	parseChartSpace,
	writeChartSpace,
	type ChartSpace,
} from '../../../chart/index';
import type { PptxChartData } from '../types';
import { writeNewChartPart } from './chart-new-part';

/** A new chart part reading an embedded workbook, and the rows of that workbook's sheet. */
export interface EmbeddedWorkbookChart {
	space: ChartSpace;
	rows: (string | number | null)[][];
	/** Model fields the part leaves out (as `writeNewChartPart` reports them). */
	unwritten: string[];
}

/** The chart part and workbook rows of `chart`, or the reason they cannot be written. */
export function createEmbeddedWorkbookChart(
	chart: PptxChartData,
): EmbeddedWorkbookChart | { reason: string } {
	const part = writeNewChartPart(chart);
	if (part.xml === undefined || !part.space)
		return {
			reason: `the shared chart core cannot write it (${part.unwritten.join(', ') || 'unknown'})`,
		};
	try {
		const { chartSpace, rows } = newChartWorkbookLayout(part.space, DEFAULT_CHART_SHEET_NAME);
		const xml = writeChartSpace({ ...chartSpace, externalDataRelId: 'rId1' });
		if (chartIssuesLostOnWrite(parseChartSpace(xml).issues).length > 0)
			return { reason: 'the chart part with cell references would lose content' };
		return { space: chartSpace, rows, unwritten: part.unwritten };
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return { reason: `its data cannot be laid out in a workbook (${message})` };
	}
}

/**
 * {@link createEmbeddedWorkbookChart} for the chart element `id`, throwing when it cannot be
 * written, so the save never drops the requested workbook.
 */
export function requireEmbeddedWorkbookChart(
	chart: PptxChartData,
	id: string,
): EmbeddedWorkbookChart {
	const created = createEmbeddedWorkbookChart(chart);
	if ('reason' in created)
		throw new Error(
			`Chart "${id}" asks for an embedded workbook (embedWorkbook), but ${created.reason}. ` +
				'Remove embedWorkbook to save its data as literal values.',
		);
	return created;
}
