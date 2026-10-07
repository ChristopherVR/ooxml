/** Resolved chart gradients, shared by Office chart painters. */
import type { DiagramFill, DiagramColor } from '../diagram/types';
import { sortGradientStops } from './gradient-stop-edit';
import { sigmaGradientStops } from '../color/sigma-gradient-stops';
import {
	buildRectPathGradientSvg,
	type RectPathGradientFillToRect,
} from '../diagram/rect-path-gradient';

export interface ChartGradientFill {
	type: 'linear' | 'radial';
	stops: Array<{ color: string; position: number; opacity?: number }>;
	angle?: number;
	scaled?: boolean;
	interpolation?: 'sigma-gamma22';
	focalPoint?: { x: number; y: number };
	path?: string;
	fillToRect?: RectPathGradientFillToRect;
}
export interface ChartSvgGradientStop {
	offset: number;
	color: string;
	opacity?: number;
}
export type ChartSvgGradientDef =
	| { kind: 'rectPath'; id: string; href: string; stops: ChartSvgGradientStop[] }
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
	const stops = fill.stops.flatMap((stop) => {
		const color = resolve(stop.color);
		return color ? [{ position: stop.position, color: color.hex, opacity: color.alpha }] : [];
	});
	return {
		type: fill.path ? 'radial' : 'linear',
		stops,
		...(fill.angle === undefined ? {} : { angle: fill.angle }),
		...(fill.scaled === undefined ? {} : { scaled: fill.scaled }),
		...(fill.path === undefined ? {} : { path: fill.path }),
		...(focus === undefined ? {} : { fillToRect: { ...focus } }),
		// Native opaque endpoint-pair linear and rectangular profiles share this curve.
		...(((!fill.path && fill.scaled === true) || fill.path === 'rect') &&
		fill.stops.length === 2 &&
		stops.length === 2 &&
		stops.every((stop) => stop.opacity === 1) &&
		stops.some((stop) => stop.position === 0) &&
		stops.some((stop) => stop.position === 100)
			? { interpolation: 'sigma-gamma22' as const }
			: {}),
		...(focus
			? { focalPoint: { x: (1 + focus.l - focus.r) / 2, y: (1 + focus.t - focus.b) / 2 } }
			: {}),
	};
}

/** Extracted from PowerPoint's COM-verified chart gradient painter. */
export function buildChartGradientDef(id: string, fill: ChartGradientFill): ChartSvgGradientDef {
	const sourceStops = sortGradientStops(fill.stops).map((stop) => ({
		offset: Math.min(Math.max(stop.position / 100, 0), 1),
		color: stop.color,
		...(stop.opacity !== undefined ? { opacity: stop.opacity } : {}),
	}));
	const stops =
		fill.interpolation === 'sigma-gamma22'
			? (sigmaGradientStops(sourceStops) ?? sourceStops)
			: sourceStops;
	if (fill.type === 'radial' && fill.path === 'rect') {
		const markup = buildRectPathGradientSvg(
			stops.map((stop) => ({
				position: stop.offset * 100,
				color: stop.color,
				...(stop.opacity === undefined ? {} : { opacity: stop.opacity }),
			})),
			undefined,
			fill.fillToRect,
			{ bands: 1024, independentOpacity: true },
		);
		return {
			kind: 'rectPath',
			id,
			href: `data:image/svg+xml,${encodeURIComponent(markup)}`,
			stops,
		};
	}
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
	const span = fill.scaled ? Math.abs(Math.cos(rad)) + Math.abs(Math.sin(rad)) : 1;
	const dx = (Math.cos(rad) * span) / 2;
	const dy = (Math.sin(rad) * span) / 2;
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
