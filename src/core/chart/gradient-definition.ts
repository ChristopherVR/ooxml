/** Resolved chart gradients, shared by Office chart painters. */
import type { DiagramFill, DiagramColor } from '../diagram/types';

export interface ChartGradientFill {
	type: 'linear' | 'radial';
	stops: Array<{ color: string; position: number; opacity?: number }>;
	angle?: number;
	focalPoint?: { x: number; y: number };
}
export interface ChartSvgGradientStop {
	offset: number;
	color: string;
	opacity?: number;
}
export type ChartSvgGradientDef =
	| {
			kind: 'linearGradient';
			id: string;
			x1: number;
			y1: number;
			x2: number;
			y2: number;
			stops: ChartSvgGradientStop[];
	  }
	| {
			kind: 'radialGradient';
			id: string;
			cx: number;
			cy: number;
			r: number;
			stops: ChartSvgGradientStop[];
	  };

/** Resolve imported DrawingML gradient stops once for every Office chart painter. */
export function resolveChartGradient(
	fill: Extract<DiagramFill, { kind: 'gradient' }>,
	resolve: (color: DiagramColor) => { hex: string; alpha: number } | undefined,
): ChartGradientFill {
	const focus = fill.fillToRect;
	return {
		type: fill.path ? 'radial' : 'linear',
		stops: fill.stops.flatMap((stop) => {
			const color = resolve(stop.color);
			return color ? [{ position: stop.position, color: color.hex, opacity: color.alpha }] : [];
		}),
		...(fill.angle === undefined ? {} : { angle: fill.angle }),
		...(focus
			? { focalPoint: { x: (1 + focus.l - focus.r) / 2, y: (1 + focus.t - focus.b) / 2 } }
			: {}),
	};
}

/** Extracted from PowerPoint's COM-verified chart gradient painter. */
export function buildChartGradientDef(id: string, fill: ChartGradientFill): ChartSvgGradientDef {
	const stops = fill.stops.map((stop) => ({
		offset: Math.min(Math.max(stop.position / 100, 0), 1),
		color: stop.color,
		...(stop.opacity !== undefined ? { opacity: stop.opacity } : {}),
	}));
	if (fill.type === 'radial') {
		const cx = fill.focalPoint?.x ?? 0.5;
		const cy = fill.focalPoint?.y ?? 0.5;
		return {
			kind: 'radialGradient',
			id,
			cx,
			cy,
			r: Math.max(...[0, 1].flatMap((x) => [0, 1].map((y) => Math.hypot(x - cx, y - cy)))),
			stops,
		};
	}
	const rad = ((fill.angle ?? 0) * Math.PI) / 180;
	const dx = Math.cos(rad) / 2;
	const dy = Math.sin(rad) / 2;
	const round = (value: number) => Math.round(value * 10000) / 10000;
	return {
		kind: 'linearGradient',
		id,
		x1: round(0.5 - dx),
		y1: round(0.5 - dy),
		x2: round(0.5 + dx),
		y2: round(0.5 + dy),
		stops,
	};
}
