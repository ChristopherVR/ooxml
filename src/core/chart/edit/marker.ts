import type { ChartSpace } from '../model';
import type { ChartMarker } from '../model-series';
import { chartSeriesEntryAt, editChartSpace } from './draft';

export type ChartMarkerEdit = { [K in keyof ChartMarker]?: ChartMarker[K] | null };

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
	edit: ChartMarkerEdit | null | undefined,
): ChartSpace {
	if (!chartSeriesMarkersSupported(space, seriesIndex))
		throw new RangeError('This chart family has no editable markers');
	if (edit?.symbol != null && !CHART_MARKER_SYMBOLS.some((symbol) => symbol === edit.symbol))
		throw new RangeError('Unsupported marker symbol');
	if (edit?.size != null && (!Number.isInteger(edit.size) || edit.size < 2 || edit.size > 72))
		throw new RangeError('Marker size must be between 2 and 72 points');
	return editChartSpace(space, (draft) => {
		const { series } = chartSeriesEntryAt(draft, seriesIndex);
		if (edit === undefined) {
			delete series.marker;
			return;
		}
		series.marker = { ...series.marker };
		if (edit === null) series.marker.symbol = 'none';
		else
			for (const key of ['symbol', 'size', 'spPr', 'extLst'] as const) {
				if (edit[key] === null) delete series.marker[key];
				else if (edit[key] !== undefined)
					Object.assign(series.marker, { [key]: structuredClone(edit[key]) });
			}
	});
}
