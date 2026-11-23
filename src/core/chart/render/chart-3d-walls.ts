// The floor, walls, gridlines and axis labels of a 3-D cartesian chart. The back wall is the plane
// at the far depth and the side wall the one on the far side of the viewer's x position, so they
// always show their inside to the camera.
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import type { ChartGradientFill } from '../gradient-definition';
import type { ChartSurface } from '../model-chrome';
import type { Chart3DCamera } from './chart-3d-camera';
import type { Chart3DLayout } from './chart-3d-layout';
import type { Vec3 } from './chart-3d-mesh';
import { strokeAttributes } from './chart-svg-analysis';
import { axisTickText } from './chart-svg-axes';
import { AXIS_COLOR, esc, GRID_COLOR, n, text, textWidth } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';
import { strokeOf, type ChartStrokeView } from './chart-view-analysis';
import type { Chart3DSurfacesSummary } from './summary';

const LABEL_SIZE = 10;

/** Resolved fill and outline for a 3-D floor or wall (`c:floor`, `c:sideWall`, `c:backWall`). */
export interface Chart3DSurfaceStyle {
	fill?: string;
	gradient?: ChartGradientFill;
	stroke?: ChartStrokeView;
}

/** Resolved 3-D floor and wall styles on a {@link ChartViewModel}. */
export interface Chart3DSurfacesView {
	floor?: Chart3DSurfaceStyle;
	sideWall?: Chart3DSurfaceStyle;
	backWall?: Chart3DSurfaceStyle;
}

function resolveSurface(
	surface: ChartSurface | undefined,
	drawingColor: (color: DrawingColor | undefined) => string | undefined,
	gradient: (fill: DrawingFill | undefined) => ChartGradientFill | undefined,
): Chart3DSurfaceStyle | undefined {
	const spPr = surface?.spPr;
	if (!spPr) return undefined;
	const style: Chart3DSurfaceStyle = {};
	if (spPr.fill?.kind === 'none') style.fill = 'none';
	else if (spPr.fill?.kind === 'solid') {
		const color = drawingColor(spPr.fill.color);
		if (color) style.fill = color;
	} else if (spPr.fill?.kind === 'gradient') {
		const resolved = gradient(spPr.fill);
		if (resolved) style.gradient = resolved;
	}
	if (spPr.line) {
		style.stroke = strokeOf(spPr.line, { color: AXIS_COLOR, width: 1 }, drawingColor) ?? {
			color: 'none',
			width: 0,
		};
	}
	return style.fill !== undefined || style.gradient !== undefined || style.stroke !== undefined
		? style
		: undefined;
}

/** Resolves `c:floor`, `c:sideWall` and `c:backWall` shape properties against the theme. */
export function resolveSurfaces3D(
	surfaces: Chart3DSurfacesSummary | undefined,
	drawingColor: (color: DrawingColor | undefined) => string | undefined,
	gradient: (fill: DrawingFill | undefined) => ChartGradientFill | undefined,
): Chart3DSurfacesView | undefined {
	if (!surfaces) return undefined;
	const floor = resolveSurface(surfaces.floor, drawingColor, gradient);
	const sideWall = resolveSurface(surfaces.sideWall, drawingColor, gradient);
	const backWall = resolveSurface(surfaces.backWall, drawingColor, gradient);
	if (!floor && !sideWall && !backWall) return undefined;
	return {
		...(floor ? { floor } : {}),
		...(sideWall ? { sideWall } : {}),
		...(backWall ? { backWall } : {}),
	};
}

function line(camera: Chart3DCamera, a: Vec3, b: Vec3, color: string): string {
	const p = camera.project(a);
	const q = camera.project(b);
	return `<line x1="${n(p.x)}" y1="${n(p.y)}" x2="${n(q.x)}" y2="${n(q.y)}" stroke="${color}" stroke-width="1"/>`;
}

function wallPolygon(
	camera: Chart3DCamera,
	pts: Vec3[],
	defaultFill: string,
	part: 'floor' | 'backWall' | 'sideWall',
	style: Chart3DSurfaceStyle | undefined,
): string {
	const points = pts
		.map((p) => {
			const q = camera.project(p);
			return `${n(q.x)},${n(q.y)}`;
		})
		.join(' ');
	const fill = style?.fill ?? defaultFill;
	const stroke = style?.stroke
		? style.stroke.color === 'none'
			? 'stroke="none"'
			: strokeAttributes(style.stroke)
		: `stroke="${AXIS_COLOR}" stroke-width="1"`;
	return `<polygon points="${points}" fill="${esc(fill)}" ${stroke} data-chart-part="${part}"/>`;
}

/** The pixel room the axis labels of a 3-D chart need left of and below its box. */
export function chart3DLabelMargins(model: ChartViewModel): { left: number; bottom: number } {
	const scale = model.valueAxis;
	const value = scale
		? Math.max(
				...scale.ticks.map((t) =>
					textWidth(axisTickText(t, model.axes?.value, scale.percent), LABEL_SIZE),
				),
				8,
			) + 6
		: 8;
	const category = Math.max(...model.categories.map((c) => textWidth(c, LABEL_SIZE)), 8) + 6;
	return model.horizontal
		? { left: Math.min(category, 120), bottom: LABEL_SIZE + 8 }
		: { left: Math.min(value, 80), bottom: LABEL_SIZE + 8 };
}

/** Floor, walls and gridlines, drawn before the marks. */
export function chart3DWallsSvg(
	model: ChartViewModel,
	layout: Chart3DLayout,
	camera: Chart3DCamera,
): string {
	const { w, h, d } = layout.size;
	const sx = camera.sideWallAtZero ? 0 : w;
	const surfaces = model.surfaces3D;
	const out: string[] = [];
	out.push(
		wallPolygon(
			camera,
			[
				[0, 0, 0],
				[w, 0, 0],
				[w, 0, d],
				[0, 0, d],
			],
			'#F2F2F2',
			'floor',
			surfaces?.floor,
		),
	);
	out.push(
		wallPolygon(
			camera,
			[
				[0, 0, d],
				[w, 0, d],
				[w, h, d],
				[0, h, d],
			],
			'none',
			'backWall',
			surfaces?.backWall,
		),
	);
	out.push(
		wallPolygon(
			camera,
			[
				[sx, 0, 0],
				[sx, 0, d],
				[sx, h, d],
				[sx, h, 0],
			],
			'none',
			'sideWall',
			surfaces?.sideWall,
		),
	);
	for (const tick of model.valueAxis?.ticks ?? []) {
		const p = layout.value(tick);
		if (layout.horizontal) {
			out.push(line(camera, [p, 0, d], [p, h, d], GRID_COLOR));
			out.push(line(camera, [p, 0, 0], [p, 0, d], GRID_COLOR));
		} else {
			out.push(line(camera, [0, p, d], [w, p, d], GRID_COLOR));
			out.push(line(camera, [sx, p, 0], [sx, p, d], GRID_COLOR));
		}
	}
	return out.join('');
}

/** Value and category labels along the front edges of the box. */
export function chart3DLabelsSvg(
	model: ChartViewModel,
	layout: Chart3DLayout,
	camera: Chart3DCamera,
	catPos: (i: number) => number,
): string {
	const { w } = layout.size;
	const sx = camera.sideWallAtZero ? 0 : w;
	const scale = model.valueAxis;
	const out: string[] = [];
	const anchor = sx === 0 ? 'end' : 'start';
	const dx = sx === 0 ? -4 : 4;
	for (const tick of scale?.ticks ?? []) {
		const label = axisTickText(tick, model.axes?.value, scale?.percent ?? false);
		const p = layout.value(tick);
		if (layout.horizontal) {
			const q = camera.project([p, 0, 0]);
			out.push(text(q.x, q.y + LABEL_SIZE + 3, label, { anchor: 'middle', size: LABEL_SIZE }));
		} else {
			const q = camera.project([sx, p, 0]);
			out.push(text(q.x + dx, q.y + 3, label, { anchor, size: LABEL_SIZE }));
		}
	}
	const a = camera.project([0, 0, 0]);
	const b = camera.project(layout.horizontal ? [0, layout.band, 0] : [layout.band, 0, 0]);
	const bandPx = Math.max(1, layout.horizontal ? Math.abs(b.y - a.y) : Math.abs(b.x - a.x));
	const extent = layout.horizontal
		? LABEL_SIZE + 2
		: Math.max(...model.categories.map((c) => textWidth(c, LABEL_SIZE)), 1);
	const stride = Math.max(1, Math.ceil(extent / bandPx));
	model.categories.forEach((label, i) => {
		if (i % stride) return;
		if (layout.horizontal) {
			const q = camera.project([sx, catPos(i), 0]);
			out.push(text(q.x + dx, q.y + 3, label, { anchor, size: LABEL_SIZE }));
		} else {
			const q = camera.project([catPos(i), 0, 0]);
			out.push(text(q.x, q.y + LABEL_SIZE + 3, label, { anchor: 'middle', size: LABEL_SIZE }));
		}
	});
	return out.join('');
}
