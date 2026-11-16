/**
 * A series' own number format (`PptxChartSeries.numberFormat`) saved through the shared chart core
 * (docs/agnostic-core-plan.md, chart step C3). The field is read from the series' label number
 * format (`c:ser/c:dLbls/c:numFmt`), falling back to the value cache's format code; the
 * object-tree writers never wrote it, so the edit was lost on save. It is written as the series'
 * label number format, not linked to the source (`setChartDataLabels` on the series), and a
 * series without labels of its own takes its group's switches so what it shows does not change.
 * Clearing the format is left unclaimed: the cache's code would read back in its place.
 *
 * @module utils/chart-neutral-save-series-format
 */

import { setChartDataLabels } from '../../../chart/index';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

export function planNeutralSeriesNumberFormat(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	if (before.series.length !== probe.series.length) return;
	probe.series.forEach((series, index) => {
		const was = before.series[index]!.numberFormat;
		const now = series.numberFormat;
		if (was === now || now === undefined || now === '') return;
		edits.push((space) => setChartDataLabels(space, { numberFormat: now }, { series: index }));
		if (was === undefined) delete series.numberFormat;
		else series.numberFormat = was;
	});
}
