import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import type { ChartMarker, ChartPlotGroup, ChartSpaceSeries } from '../model';
import { automaticMarkerSymbol } from '../marker-shape';

/** Marker declarations retained by the neutral summary, including point overrides. */
export interface ChartMarkerSummary {
	marker?: ChartMarker;
	pointMarkers?: Record<number, ChartMarker>;
}

/** Marker paint resolved against a host's chart theme. Sizes remain point diameters. */
export interface ChartMarkerView {
	symbol?: string;
	size?: number;
	fill?: string;
	stroke?: string;
	strokeWidth?: number;
}

export interface ChartMarkerSeriesView {
	color: string;
	pointColors?: string[];
	shadowFilter?: string;
	marker?: ChartMarkerView;
	pointMarkers?: Record<number, ChartMarkerView>;
}

/** Carries OOXML marker defaults into each series, before host/theme resolution. */
export function summaryMarkers(
	source: ChartSpaceSeries,
	group: ChartPlotGroup,
	seriesPosition: number,
): ChartMarkerSummary {
	if (!['line', 'scatter', 'radar'].includes(group.kind)) return {};
	const enabled =
		group.kind === 'line'
			? group.marker === true
			: group.kind === 'radar'
				? group.radarStyle === 'marker'
				: !['line', 'smooth', 'none'].includes(group.scatterStyle ?? 'marker');
	const symbol = source.marker?.symbol;
	const marker: ChartMarker = {
		...source.marker,
		symbol:
			symbol && symbol !== 'auto'
				? symbol
				: source.marker || enabled
					? automaticMarkerSymbol(source.index ?? seriesPosition)
					: 'none',
	};
	const pointMarkers: Record<number, ChartMarker> = {};
	for (const point of source.dataPoints) {
		if (point.index === undefined || !point.marker) continue;
		const override = point.marker;
		pointMarkers[point.index] = {
			...marker,
			...override,
			...(override.symbol === 'auto' || (!source.marker && override.symbol === undefined)
				? { symbol: automaticMarkerSymbol(source.index ?? seriesPosition) }
				: {}),
			...(marker.spPr || override.spPr
				? {
						spPr: {
							...marker.spPr,
							...override.spPr,
							...(marker.spPr?.line || override.spPr?.line
								? { line: { ...marker.spPr?.line, ...override.spPr?.line } }
								: {}),
						},
					}
				: {}),
		};
	}
	return { marker, ...(Object.keys(pointMarkers).length ? { pointMarkers } : {}) };
}

/** Resolves marker DrawingML paint without depending on a product model. */
export function seriesMarkerViews(
	source: ChartMarkerSummary,
	drawingColor: (color: DrawingColor | undefined) => string | undefined,
): Pick<ChartMarkerSeriesView, 'marker' | 'pointMarkers'> {
	const fillColor = (fill: DrawingFill | undefined): string | undefined =>
		fill?.kind === 'none' ? 'none' : fill?.kind === 'solid' ? drawingColor(fill.color) : undefined;
	const resolve = (marker: ChartMarker): ChartMarkerView => {
		const fill = fillColor(marker.spPr?.fill);
		const stroke = fillColor(marker.spPr?.line?.fill);
		const width = marker.spPr?.line?.widthEmu;
		return {
			...(marker.symbol !== undefined ? { symbol: marker.symbol } : {}),
			...(marker.size !== undefined ? { size: marker.size } : {}),
			...(fill !== undefined ? { fill } : {}),
			...(stroke !== undefined ? { stroke } : {}),
			...(width !== undefined ? { strokeWidth: width / 9525 } : {}),
		};
	};
	return {
		...(source.marker ? { marker: resolve(source.marker) } : {}),
		...(source.pointMarkers
			? {
					pointMarkers: Object.fromEntries(
						Object.entries(source.pointMarkers).map(([i, marker]) => [i, resolve(marker)]),
					),
				}
			: {}),
	};
}
