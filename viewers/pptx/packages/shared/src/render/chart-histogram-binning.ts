/**
 * Compatibility entry: histogram binning lives in ooxml-core/chart
 * (`histogram-binning.ts`). This module only binds the viewer's axis number
 * formatting to the bin labels, which the core takes as a parameter.
 *
 * @module chart-histogram-binning
 */
import type { PptxChartHistogramOptions } from 'pptx-viewer-core';
import { computeHistogramBins as computeCoreHistogramBins } from 'pptx-viewer-core/chart';
import type { HistogramBin } from 'pptx-viewer-core/chart';

import { formatAxisValue } from './chart-view-model';

export { aggregateByCategory, scottBinCount } from 'pptx-viewer-core/chart';
export type { HistogramBin } from 'pptx-viewer-core/chart';

/** Bins raw observations, labelling the bin edges with the viewer's axis formatting. */
export function computeHistogramBins(
	values: ReadonlyArray<number>,
	options: PptxChartHistogramOptions,
): HistogramBin[] {
	return computeCoreHistogramBins(values, options, formatAxisValue);
}
