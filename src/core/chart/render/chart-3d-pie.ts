// A 3-D pie: the disc seen at elevation `rotX` is an ellipse (minor / major = sin rotX), thickened
// by an extrusion whose height follows `depthPercent` and the tilt. Slices start at twelve
// o'clock (plus `firstSliceAngle`) and run clockwise like the flat pie. Exploded slices shift
// outward along their bisector, expose their radial cut walls and are depth-sorted back-to-front.
import { resolveChart3DProjection, type Chart3DViewInput } from './chart-3d-projection';
import { shadeChart3DFace } from './chart-3d-shading';
import { strokeAttributes } from './chart-svg-analysis';
import { draw, hasOutsideLabels, measure } from './chart-svg-labels';
import { n, type Rect } from './chart-svg-util';
import type { ChartSeriesView, ChartViewModel } from './chart-view';

const STEP = Math.PI / 45;
const LABEL_GAP = 4;

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

const pt = (cx: number, cy: number, rx: number, ry: number, a: number, dy: number): string =>
	`${n(cx + rx * Math.cos(a))},${n(cy + ry * Math.sin(a) + dy)}`;

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

interface PieSlice3D {
	index: number;
	a0: number;
	a1: number;
	mid: number;
	sweep: number;
	scx: number;
	scy: number;
	color: string;
	wallColor: string;
	cutColor: string;
	filter: string;
	walls: string[];
	top: string;
}

function cutWallPolygon(
	cx: number,
	cy: number,
	rx: number,
	ry: number,
	thickness: number,
	angle: number,
	color: string,
	index: number,
	filter: string,
): string {
	const ex = cx + rx * Math.cos(angle);
	const ey = cy + ry * Math.sin(angle);
	const pts = [
		`${n(cx)},${n(cy)}`,
		`${n(ex)},${n(ey)}`,
		`${n(ex)},${n(ey + thickness)}`,
		`${n(cx)},${n(cy + thickness)}`,
	].join(' ');
	return `<polygon points="${pts}" fill="${color}" stroke="${color}" stroke-width="0.5" data-chart-series="0" data-chart-point="${index}"${filter}/>`;
}

function drawPie3DLabels(
	model: ChartViewModel,
	series: ChartSeriesView,
	slices: readonly PieSlice3D[],
	rx: number,
	ry: number,
	out: string[],
): void {
	if (!series.labels) return;
	for (const slice of slices) {
		const label = series.labels[slice.index];
		if (!label) continue;
		const m = measure(model, label);
		const ux = Math.cos(slice.mid);
		const uy = Math.sin(slice.mid);
		const position = label.position ?? 'bestFit';
		const outside = position === 'outEnd' || (position === 'bestFit' && slice.sweep < 0.35);
		let lx: number;
		let ly: number;
		if (position === 'ctr') {
			lx = slice.scx + rx * 0.55 * ux;
			ly = slice.scy + ry * 0.55 * uy;
		} else if (position === 'inEnd') {
			lx = slice.scx + rx * 0.78 * ux;
			ly = slice.scy + ry * 0.78 * uy;
		} else if (outside) {
			const rimX = slice.scx + rx * ux;
			const rimY = slice.scy + ry * uy;
			lx = rimX + (LABEL_GAP + m.w / 2) * ux;
			ly = rimY + (LABEL_GAP + m.h / 2) * uy;
			if (label.showLeaderLines !== false) {
				const stroke = label.leaderLine ?? { color: '#94a3b8', width: 0.75 };
				out.push(
					`<line x1="${n(rimX)}" y1="${n(rimY)}" x2="${n(lx)}" y2="${n(ly)}" ${strokeAttributes(stroke)} data-chart-part="leaderLine"/>`,
				);
			}
		} else {
			lx = slice.scx + rx * 0.62 * ux;
			ly = slice.scy + ry * 0.62 * uy;
		}
		draw(m, label, lx, ly, 'middle', out);
	}
}

export function chart3DPieSvg(model: ChartViewModel, area: Rect, view: Chart3DViewInput): string {
	const series = model.series[0];
	if (!series) return '';
	const total = series.values.reduce<number>((sum, v) => sum + (v !== null && v > 0 ? v : 0), 0);
	if (total <= 0) return '';
	const margin = hasOutsideLabels(model) ? Math.min(area.w, area.h) * 0.12 : 0;
	const fitArea: Rect =
		margin > 0
			? {
					x: area.x + margin,
					y: area.y + margin,
					w: Math.max(1, area.w - margin * 2),
					h: Math.max(1, area.h - margin * 2),
				}
			: area;
	const baseGeom = chart3DPieGeometry(fitArea, view);
	const explosions = series.values.map((_, i) =>
		Math.max(0, Math.min(400, series.pointExplosions?.[i] ?? series.explosion ?? 0)),
	);
	const maxExplosion = Math.max(0, ...explosions) / 100;
	const scale = maxExplosion > 0 ? 1 / (1 + maxExplosion) : 1;
	const g: Chart3DPieGeometry = {
		cx: baseGeom.cx,
		cy: baseGeom.cy,
		rx: baseGeom.rx * scale,
		ry: baseGeom.ry * scale,
		thickness: baseGeom.thickness * scale,
	};
	const slices: PieSlice3D[] = [];
	let angle = -Math.PI / 2 + ((model.firstSliceAngle ?? 0) * Math.PI) / 180;
	series.values.forEach((v, i) => {
		if (v === null || v <= 0) return;
		const sweep = (v / total) * Math.PI * 2;
		const a0 = angle;
		const a1 = angle + sweep;
		angle = a1;
		const mid = (a0 + a1) / 2;
		const exp = (explosions[i] ?? 0) / 100;
		const scx = g.cx + g.rx * exp * Math.cos(mid);
		const scy = g.cy + g.ry * exp * Math.sin(mid);
		const color = series.pointColors?.[i] ?? series.color;
		const wallColor = shadeChart3DFace(color, 'right');
		const cutColor = shadeChart3DFace(color, 'left');
		const filter = series.shadowFilter ? ` filter="${series.shadowFilter}"` : '';
		const walls: string[] = [];
		const full = sweep >= Math.PI * 2 - 1e-6;
		if (maxExplosion > 0 && !full && g.thickness > 1e-3) {
			if (Math.cos(a0) < -1e-6)
				walls.push(cutWallPolygon(scx, scy, g.rx, g.ry, g.thickness, a0, cutColor, i, filter));
			if (Math.cos(a1) > 1e-6)
				walls.push(cutWallPolygon(scx, scy, g.rx, g.ry, g.thickness, a1, cutColor, i, filter));
		}
		for (const [lo, hi] of frontRanges(a0, a1)) {
			const steps = Math.max(2, Math.ceil((hi - lo) / STEP));
			const upper: string[] = [];
			const lower: string[] = [];
			for (let s = 0; s <= steps; s++) {
				const a = lo + ((hi - lo) * s) / steps;
				upper.push(pt(scx, scy, g.rx, g.ry, a, 0));
				lower.push(pt(scx, scy, g.rx, g.ry, a, g.thickness));
			}
			walls.push(
				`<polygon points="${[...upper, ...lower.reverse()].join(' ')}" fill="${wallColor}" stroke="${wallColor}" stroke-width="0.5" data-chart-series="0" data-chart-point="${i}"${filter}/>`,
			);
		}
		const start = pt(scx, scy, g.rx, g.ry, a0, 0);
		const end = pt(scx, scy, g.rx, g.ry, a1, 0);
		const d = full
			? `M${start} A${n(g.rx)},${n(g.ry)} 0 1 1 ${pt(scx, scy, g.rx, g.ry, a0 + Math.PI, 0)} A${n(g.rx)},${n(g.ry)} 0 1 1 ${start} Z`
			: `M${n(scx)},${n(scy)} L${start} A${n(g.rx)},${n(g.ry)} 0 ${sweep > Math.PI ? 1 : 0} 1 ${end} Z`;
		const top = `<path d="${d}" fill="${color}" stroke="#FFFFFF" stroke-width="1" data-chart-series="0" data-chart-point="${i}"${filter}/>`;
		slices.push({
			index: i,
			a0,
			a1,
			mid,
			sweep,
			scx,
			scy,
			color,
			wallColor,
			cutColor,
			filter,
			walls,
			top,
		});
	});
	const out: string[] = [];
	if (maxExplosion > 0) {
		const ordered = [...slices].sort((a, b) => Math.sin(a.mid) - Math.sin(b.mid));
		for (const s of ordered) {
			out.push(...s.walls, s.top);
		}
	} else {
		for (const s of slices) out.push(...s.walls);
		for (const s of slices) out.push(s.top);
	}
	drawPie3DLabels(model, series, slices, g.rx, g.ry, out);
	return out.join('');
}
