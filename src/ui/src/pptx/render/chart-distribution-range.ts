import type { PptxChartSeries } from 'ooxml-core/pptx';

import type { ValueRange } from './chart-view-model';

/** Range with zero included and the upper bound floored at one. */
export function distributionRange(series: ReadonlyArray<PptxChartSeries>): ValueRange {
	const all = series.flatMap((item) => item.values);
	const min = Math.min(...all, 0);
	const max = Math.max(...all, 1);
	return { min, max, span: Math.max(max - min, 1) };
}
