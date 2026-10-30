// A 3-D pie: the disc seen at elevation `rotX` is an ellipse (minor / major = sin rotX), thickened
// by an extrusion whose height follows `depthPercent` and the tilt. Slices start at twelve
// o'clock and run clockwise like the flat pie. The visible part of each slice wall is the front
// half of its rim; the tops are painted over the walls.
import { resolveChart3DProjection, type Chart3DViewInput } from './chart-3d-projection';
import { shadeChart3DFace } from './chart-3d-shading';
import { n, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

const STEP = Math.PI / 45;

export interface Chart3DPieGeometry {
	cx: number;
	cy: number;
	/** Horizontal and vertical radius of the top face. */
	rx: number;
	ry: number;
	/** Height of the extrusion in pixels. */
	thickness: number;
}

/** The pie ellipse and thickness that fit `area` for a tilt and depth. */
export function chart3DPieGeometry(area: Rect, view: Chart3DViewInput): Chart3DPieGeometry {
	const tilt = resolveChart3DProjection('pie3D', view).rotXDeg;
	const sin = Math.min(1, Math.max(0.08, Math.sin((Math.abs(tilt) * Math.PI) / 180)));
	const cos = Math.cos((Math.min(90, Math.abs(tilt)) * Math.PI) / 180);
	const depth = Math.min(2000, Math.max(20, view.depthPercent ?? 100)) / 100;
	const rx0 = Math.max(1, Math.min(area.w / 2, area.h / 2 / sin) - 4);
	const thickness = rx0 * 0.35 * cos * depth;
	const rx = Math.min(rx0, Math.max(1, (area.h - thickness) / 2 / sin));
	const ry = rx * sin;
	const thick = Math.min(thickness, Math.max(0, area.h - ry * 2));
	return {
		cx: area.x + area.w / 2,
		cy: area.y + (area.h - thick) / 2 + ry,
		rx,
		ry,
		thickness: thick,
	};
}

const pt = (g: Chart3DPieGeometry, a: number, dy: number): string =>
	`${n(g.cx + g.rx * Math.cos(a))},${n(g.cy + g.ry * Math.sin(a) + dy)}`;

/** The front-facing parts of the rim between `a0` and `a1` (angles in radians, screen y down). */
function frontRanges(a0: number, a1: number): [number, number][] {
	const out: [number, number][] = [];
	for (let k = Math.floor(a0 / (2 * Math.PI)) - 1; k <= Math.ceil(a1 / (2 * Math.PI)); k++) {
		const lo = Math.max(a0, k * 2 * Math.PI);
		const hi = Math.min(a1, k * 2 * Math.PI + Math.PI);
		if (hi - lo > 1e-6) out.push([lo, hi]);
	}
	return out;
}

export function chart3DPieSvg(model: ChartViewModel, area: Rect, view: Chart3DViewInput): string {
	const series = model.series[0];
	if (!series) return '';
	const total = series.values.reduce<number>((sum, v) => sum + (v !== null && v > 0 ? v : 0), 0);
	if (total <= 0) return '';
	const g = chart3DPieGeometry(area, view);
	const walls: string[] = [];
	const tops: string[] = [];
	let angle = -Math.PI / 2;
	series.values.forEach((v, i) => {
		if (v === null || v <= 0) return;
		const sweep = (v / total) * Math.PI * 2;
		const a0 = angle;
		const a1 = angle + sweep;
		angle = a1;
		const color = series.pointColors?.[i] ?? series.color;
		const wall = shadeChart3DFace(color, 'right');
		const top = color;
		const filter = series.shadowFilter ? ` filter="${series.shadowFilter}"` : '';
		for (const [lo, hi] of frontRanges(a0, a1)) {
			const steps = Math.max(2, Math.ceil((hi - lo) / STEP));
			const upper: string[] = [];
			const lower: string[] = [];
			for (let s = 0; s <= steps; s++) {
				const a = lo + ((hi - lo) * s) / steps;
				upper.push(pt(g, a, 0));
				lower.push(pt(g, a, g.thickness));
			}
			walls.push(
				`<polygon points="${[...upper, ...lower.reverse()].join(' ')}" fill="${wall}" stroke="${wall}" stroke-width="0.5" data-chart-series="0" data-chart-point="${i}"${filter}/>`,
			);
		}
		const full = sweep >= Math.PI * 2 - 1e-6;
		const start = pt(g, a0, 0);
		const end = pt(g, a1, 0);
		const d = full
			? `M${pt(g, a0, 0)} A${n(g.rx)},${n(g.ry)} 0 1 1 ${pt(g, a0 + Math.PI, 0)} A${n(g.rx)},${n(g.ry)} 0 1 1 ${start} Z`
			: `M${n(g.cx)},${n(g.cy)} L${start} A${n(g.rx)},${n(g.ry)} 0 ${sweep > Math.PI ? 1 : 0} 1 ${end} Z`;
		tops.push(
			`<path d="${d}" fill="${top}" stroke="#FFFFFF" stroke-width="1" data-chart-series="0" data-chart-point="${i}"${filter}/>`,
		);
	});
	return walls.join('') + tops.join('');
}
