/**
 * Write chart-edited values back into an embedded xlsx workbook
 * (`ppt/embeddings/*.xlsx`).
 *
 * The writer itself moved to the format-neutral chart area
 * (`chart/workbook/writer.ts`, shared with Word); this module keeps the pptx
 * names and the object-tree reader that collects one write from a chart cache
 * container.
 *
 * @module chart-xlsx-writer
 */

import type { ChartWorkbookWrite } from '../../../chart/workbook/writer';
import type { IPptxXmlLookupService } from '../services';
import type { XmlObject } from '../types';

export {
	writeChartWorkbookUpdates,
	type ChartWorkbookWriteResult,
} from '../../../chart/workbook/writer';

/** One chart data reference (series name, category, or value list) to write back. */
export type PptxChartWorkbookWrite = ChartWorkbookWrite;

/**
 * Build the `c:f` formula write for one chart cache container (a series'
 * `c:tx`, `c:cat`/`c:xVal`, or `c:val`/`c:yVal`), given the values already
 * written to its cache. Returns `undefined` when the container has no
 * `numRef`/`strRef` formula reference - e.g. a brand-new series built
 * without one, or a `numLit`/`strLit` inline literal - since there is then
 * no linked worksheet range to update.
 */
export function collectChartWorkbookWrite(
	xmlLookupService: IPptxXmlLookupService,
	container: XmlObject | undefined,
	isNumeric: boolean,
	values: string[],
): PptxChartWorkbookWrite | undefined {
	if (!container) {
		return undefined;
	}
	const refNode = xmlLookupService.getChildByLocalName(container, isNumeric ? 'numRef' : 'strRef');
	const formula = refNode ? xmlLookupService.getScalarChildByLocalName(refNode, 'f') : undefined;
	return formula ? { formula, isNumeric, values } : undefined;
}
