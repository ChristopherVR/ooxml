import { resolveDrawingColor } from '../../../drawingml/drawing-color';
import type { DrawingColor, DrawingLine } from '../../../drawingml/types';
import type { PptxChartLineStyle } from '../types';
import type { ChartTreeReader } from './chart-series-tree-fields';

/** Resolve neutral colors using the host only for its theme aliases and color-map override. */
export function chartColorFromNeutral(
	color: DrawingColor | undefined,
	reader: Pick<ChartTreeReader, 'parseColor'>,
): string | undefined {
	if (!color) return undefined;
	const resolved = resolveDrawingColor(color, {
		scheme: (name) => reader.parseColor({ 'a:schemeClr': { '@_val': name } }),
	});
	return resolved?.hex;
}

/** The compatibility line descriptor; authored tokens and unsupported details remain in core. */
export function chartLineFromNeutral(
	line: DrawingLine | undefined,
	reader: Pick<ChartTreeReader, 'parseColor'>,
): PptxChartLineStyle {
	const result: PptxChartLineStyle = {};
	const color =
		line?.fill?.kind === 'solid' ? chartColorFromNeutral(line.fill.color, reader) : undefined;
	if (color) result.color = color;
	if (line?.widthEmu !== undefined && Number.isFinite(line.widthEmu))
		result.width = line.widthEmu / 12700;
	if (line?.dash) result.dashStyle = line.dash;
	return result;
}

/** Project the shared gradient model into the compatibility painter's descriptor. */
export function chartGradientFromNeutral(
	fill: import('../../../drawingml/types').DrawingFill | undefined,
	reader: Pick<ChartTreeReader, 'parseColor'>,
): import('../types').PptxChartGradientFill | undefined {
	if (fill?.kind !== 'gradient') return undefined;
	const stops = fill.stops.flatMap((stop) => {
		const resolved = resolveDrawingColor(stop.color, {
			scheme: (name) => reader.parseColor({ 'a:schemeClr': { '@_val': name } }),
		});
		return resolved
			? [
					{
						color: resolved.hex,
						position: stop.position,
						...(resolved.alpha < 1 ? { opacity: resolved.alpha } : {}),
					},
				]
			: [];
	});
	if (stops.length < 2) return undefined;
	if (!fill.path) return { type: 'linear', stops, angle: fill.angle ?? 0 };
	const rect = fill.fillToRect;
	const clamp = (value: number) => Math.max(0, Math.min(1, value));
	return {
		type: 'radial',
		stops,
		...(rect
			? { focalPoint: { x: clamp((rect.l + 1 - rect.r) / 2), y: clamp((rect.t + 1 - rect.b) / 2) } }
			: {}),
		...(fill.path === 'circle' ? {} : { path: fill.path }),
	};
}

/** The small shape descriptor used by markers, points, walls and up/down bars. */
export function chartShapeFromNeutral(
	shape: import('../../../chart/index').ChartShapeProperties | undefined,
	reader: Pick<ChartTreeReader, 'parseColor'>,
): import('../types').PptxChartShapeProps | undefined {
	if (!shape) return undefined;
	const result: import('../types').PptxChartShapeProps = {};
	const fill =
		shape.fill?.kind === 'solid' ? chartColorFromNeutral(shape.fill.color, reader) : undefined;
	if (fill) result.fillColor = fill;
	const line = chartLineFromNeutral(shape.line, reader);
	if (shape.line?.fill?.kind === 'none') result.lineNoFill = true;
	if (line.color) result.strokeColor = line.color;
	if (line.width !== undefined) result.strokeWidth = line.width;
	if (line.dashStyle) result.strokeDashStyle = line.dashStyle;
	return Object.keys(result).length ? result : undefined;
}
