import type { PptxChartData } from '../types';

/** Only fields mirrored by classic chart workbook write-back, excluding visual edits. */
export function chartWorkbookDataSignature(data: PptxChartData): string {
	return JSON.stringify({
		categories: data.dateCategories?.values ?? data.categories,
		series: data.series.map((series) => ({ name: series.name, values: series.values })),
	});
}
