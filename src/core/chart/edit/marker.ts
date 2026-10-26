import type { ChartSpace } from '../model';
import type { ChartMarker } from '../model-series';
import { chartSeriesEntryAt, editChartSpace } from './draft';

export const CHART_MARKER_SYMBOLS = [
	'auto',
	'none',
	'circle',
	'dash',
	'diamond',
	'dot',
	'plus',
	'square',
	'star',
	'triangle',
	'x',
] as const;

export function chartSeriesMarkersSupported(space: ChartSpace, seriesIndex: number): boolean {
	const { group } = chartSeriesEntryAt(space, seriesIndex);
	return ['line', 'scatter', 'radar'].includes(group.kind) && !group.element.includes('3D');
}

/** Partial marker edits preserve authored shape properties; null explicitly suppresses inherited markers. */
export function setChartMarker(
	space: ChartSpace,
	seriesIndex: number,
	edit: Partial<ChartMarker> | null,
): ChartSpace {
	if (!chartSeriesMarkersSupported(space, seriesIndex))
		throw new RangeError('This chart family has no editable markers');
	if (edit?.symbol !== undefined && !CHART_MARKER_SYMBOLS.some((symbol) => symbol === edit.symbol))
		throw new RangeError('Unsupported marker symbol');
	if (edit?.size !== undefined && (!Number.isInteger(edit.size) || edit.size < 2 || edit.size > 72))
		throw new RangeError('Marker size must be between 2 and 72 points');
	return editChartSpace(space, (draft) => {
		const { group, series } = chartSeriesEntryAt(draft, seriesIndex);
		series.marker = {
			...series.marker,
			...(edit === null ? { symbol: 'none' } : structuredClone(edit)),
		};
		if (edit !== null && series.marker.symbol === undefined) series.marker.symbol = 'auto';
		if (group.kind === 'line') group.marker = true;
	});
}
