import { buildMarkerPrimitive, markerRadius } from '../marker-shape';
import type { ChartMarkerSeriesView } from './chart-view-markers';
import { esc, n } from './chart-svg-util';

/** Radius for label placement, including an individual point's marker size. */
export function chartMarkerRadius(
	series: ChartMarkerSeriesView,
	index: number,
	defaultRadius = 3,
): number {
	const marker = { ...series.marker, ...series.pointMarkers?.[index] };
	if (marker.symbol === 'none') return 0;
	const radius = markerRadius(marker.size, defaultRadius);
	return marker.symbol === 'dot' ? Math.max(radius * 0.6, 1) : radius;
}

/** Paints the same marker geometry as PowerPoint, with escaped SVG attributes. */
export function chartSeriesMarkerSvg(
	series: ChartMarkerSeriesView,
	index: number,
	cx: number,
	cy: number,
	defaultRadius = 3,
): string {
	const marker = { ...series.marker, ...series.pointMarkers?.[index] };
	const primitive = buildMarkerPrimitive({
		...marker,
		symbol: marker.symbol,
		size: marker.size,
		cx,
		cy,
		fill: marker.fill ?? series.pointColors?.[index] ?? series.color,
		defaultRadius,
	});
	if (!primitive) return '';
	let geometry: string;
	switch (primitive.kind) {
		case 'circle':
			geometry = `cx="${n(primitive.cx)}" cy="${n(primitive.cy)}" r="${n(primitive.r)}"`;
			break;
		case 'rect':
			geometry = `x="${n(primitive.x)}" y="${n(primitive.y)}" width="${n(primitive.w)}" height="${n(primitive.h)}"`;
			break;
		case 'path':
			geometry = `d="${esc(primitive.d)}"`;
			break;
		case 'polygon':
			geometry = `points="${esc(primitive.points)}"`;
			break;
	}
	const outline = primitive.stroke
		? ` stroke="${esc(primitive.stroke)}" stroke-width="${n(primitive.strokeWidth ?? 1)}"`
		: '';
	const filter = series.shadowFilter ? ` filter="${esc(series.shadowFilter)}"` : '';
	return `<${primitive.kind} ${geometry} fill="${esc(primitive.fill)}"${outline}${filter}/>`;
}
