import { smoothLinePath, type LinePoint } from '../line-path';
import { strokeAttributes } from './chart-svg-analysis';
import { n } from './chart-svg-util';
import type { ChartSeriesView } from './chart-view';

/** Paint one contiguous series run; missing values split runs before this boundary. */
export function chartLineSvg(points: LinePoint[], series: ChartSeriesView): string {
	if (points.length < 2) return '';
	const stroke = series.stroke;
	const attrs = strokeAttributes(
		stroke
			? { ...stroke, color: stroke.color === 'none' ? 'none' : series.color }
			: { color: series.color, width: 2.25 },
	);
	const filter = series.shadowFilter ? ` filter="${series.shadowFilter}"` : '';
	return series.smooth
		? `<path d="${smoothLinePath(points)}" fill="none" ${attrs} stroke-linejoin="round"${filter}/>`
		: `<polyline points="${points.map((p) => `${n(p.x)},${n(p.y)}`).join(' ')}" fill="none" ${attrs} stroke-linejoin="round"${filter}/>`;
}
