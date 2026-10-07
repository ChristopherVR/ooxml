/**
 * Compatibility entry: histogram binning lives in ooxml-core/chart
 * (`histogram-binning.ts`). This module only binds the viewer's axis number
 * formatting to the bin labels, which the core takes as a parameter.
 *
 * @module chart-histogram-binning
 */
import type { PptxChartHistogramOptions } from 'ooxml-core/pptx';
import { computeHistogramBins as computeCoreHistogramBins } from 'ooxml-core/chart';
import type { HistogramBin } from 'ooxml-core/chart';

import { formatAxisValue } from './chart-view-model';

export { aggregateByCategory, scottBinCount } from 'ooxml-core/chart';
export type { HistogramBin } from 'ooxml-core/chart';

/** Bins raw observations, labelling the bin edges with the viewer's axis formatting. */
export function computeHistogramBins(
	values: ReadonlyArray<number>,
	options: PptxChartHistogramOptions,
): HistogramBin[] {
	return computeCoreHistogramBins(values, options, formatAxisValue);
}
