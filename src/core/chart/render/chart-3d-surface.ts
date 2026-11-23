// Mesh generation for a 3-D surface chart and 2-D top-down contour renderer (`c:surfaceChart`):
// categories along x, series along z (or y in contour view), the cached values as heights. Every
// grid cell is coloured by the value band it falls in, shaded by slope in 3-D, or stroked without
// fill when `c:wireframe` is enabled.
import type { Chart3DMesh, Vec3 } from './chart-3d-mesh';
import type { Chart3DLayout } from './chart-3d-layout';
import { shade } from '../style-palette';
import { chartAreaRect } from './chart-svg-appearance';
import { AXIS_COLOR, n, text, textWidth, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

const LIGHT: Vec3 = (() => {
	const v: Vec3 = [-0.3, 0.8, -0.5];
	const len = Math.hypot(...v);
	return [v[0] / len, v[1] / len, v[2] / len];
})();
const HEX = /^#[0-9a-f]{6}$/i;
/** Band colours when the series carry none a surface legend could use. */
const BAND_FALLBACK = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const LABEL_SIZE = 10;

function bandColor(model: ChartViewModel, ratio: number): string {
	const scale = model.valueAxis;
	const rows = model.series.length;
	const bandCount = Math.max(1, (scale?.ticks.length ?? 2) - 1);
	const band = Math.max(0, Math.min(bandCount - 1, Math.floor(ratio * bandCount)));
	const seriesColor = model.series[band % rows]?.color;
	return seriesColor && HEX.test(seriesColor)
		? seriesColor
		: (BAND_FALLBACK[band % BAND_FALLBACK.length] ?? '#4472C4');
}

export function chart3DSurfaceMeshes(model: ChartViewModel, layout: Chart3DLayout): Chart3DMesh[] {
	const cols = model.categories.length;
	const rows = model.series.length;
	const scale = model.valueAxis;
	if (cols < 2 || rows < 1 || !scale) return [];
	const zRows = rows === 1 ? 2 : rows;
	const wireframe = model.wireframe === true;
	const height = (r: number, c: number): number =>
		layout.value(model.series[Math.min(r, rows - 1)]?.values[c] ?? 0);
	const x = (c: number): number => (c / (cols - 1)) * layout.catLength;
	const z = (r: number): number => (r / (zRows - 1)) * layout.size.d;
	const meshes: Chart3DMesh[] = [];
	for (let r = 0; r + 1 < zRows; r++)
		for (let c = 0; c + 1 < cols; c++) {
			const pts: Vec3[] = [
				[x(c), height(r, c), z(r)],
				[x(c + 1), height(r, c + 1), z(r)],
				[x(c + 1), height(r + 1, c + 1), z(r + 1)],
				[x(c), height(r + 1, c), z(r + 1)],
			];
			const avg = pts.reduce((sum, p) => sum + p[1], 0) / 4;
			const color = bandColor(model, avg / Math.max(1e-6, layout.valueLength));
			const [a, b, , d] = pts as [Vec3, Vec3, Vec3, Vec3];
			let fill: string;
			let stroke: string | undefined;
			if (wireframe) {
				fill = 'none';
				stroke = color;
			} else {
				const e1: Vec3 = [b[0] - a[0], b[1] - a[1], b[2] - a[2]];
				const e2: Vec3 = [d[0] - a[0], d[1] - a[1], d[2] - a[2]];
				const nx = e1[1] * e2[2] - e1[2] * e2[1];
				const ny = e1[2] * e2[0] - e1[0] * e2[2];
				const nz = e1[0] * e2[1] - e1[1] * e2[0];
				const len = Math.hypot(nx, ny, nz) || 1;
				const lit = Math.abs((nx * LIGHT[0] + ny * LIGHT[1] + nz * LIGHT[2]) / len);
				fill = shade(color, 0.35 * (1 - lit));
			}
			meshes.push({
				series: Math.min(r, rows - 1),
				point: c,
				color,
				center: [(a[0] + b[0]) / 2, avg, (a[2] + d[2]) / 2],
				faces: [
					{
						pts,
						normal: undefined,
						kind: 'top',
						fill,
						...(stroke !== undefined ? { stroke } : {}),
					},
				],
			});
		}
	return meshes;
}

/** Top-down 2-D contour chart (`c:surfaceChart`), filled by value band or wireframe. */
export function chartContourSvg(model: ChartViewModel, area: Rect): string {
	const out: string[] = [chartAreaRect(model, 'plotArea', area)];
	const cols = model.categories.length;
	const rows = model.series.length;
	const scale = model.valueAxis;
	if (cols < 2 || rows < 1 || !scale) return out.join('');
	const zRows = rows === 1 ? 2 : rows;
	const left = Math.min(
		80,
		Math.max(...model.series.map((s) => textWidth(s.name, LABEL_SIZE)), 8) + 8,
	);
	const bottom = LABEL_SIZE + 10;
	const plot: Rect = {
		x: area.x + left,
		y: area.y + 6,
		w: Math.max(1, area.w - left - 8),
		h: Math.max(1, area.h - bottom - 8),
	};
	const cellW = plot.w / (cols - 1);
	const cellH = plot.h / (zRows - 1);
	const span = scale.max - scale.min || 1;
	const wireframe = model.wireframe === true;
	const val = (r: number, c: number): number =>
		model.series[Math.min(r, rows - 1)]?.values[c] ?? scale.min;
	for (let r = 0; r + 1 < zRows; r++) {
		for (let c = 0; c + 1 < cols; c++) {
			const avg = (val(r, c) + val(r, c + 1) + val(r + 1, c + 1) + val(r + 1, c)) / 4;
			const ratio = (avg - scale.min) / span;
			const color = bandColor(model, ratio);
			const x = plot.x + c * cellW;
			const y = plot.y + plot.h - (r + 1) * cellH;
			const fill = wireframe ? 'none' : color;
			const strokeWidth = wireframe ? '1' : '0.5';
			out.push(
				`<rect x="${n(x)}" y="${n(y)}" width="${n(cellW)}" height="${n(cellH)}" fill="${fill}" stroke="${color}" stroke-width="${strokeWidth}" data-chart-part="contourCell" data-chart-series="${Math.min(r, rows - 1)}" data-chart-point="${c}"/>`,
			);
		}
	}
	out.push(
		`<rect x="${n(plot.x)}" y="${n(plot.y)}" width="${n(plot.w)}" height="${n(plot.h)}" fill="none" stroke="${AXIS_COLOR}" stroke-width="1"/>`,
	);
	model.categories.forEach((label, c) => {
		const x = plot.x + (c / (cols - 1)) * plot.w;
		out.push(
			text(x, plot.y + plot.h + LABEL_SIZE + 3, label, { anchor: 'middle', size: LABEL_SIZE }),
		);
	});
	model.series.forEach((s, r) => {
		const y = rows === 1 ? plot.y + plot.h / 2 : plot.y + plot.h - (r / (zRows - 1)) * plot.h;
		out.push(text(plot.x - 4, y + 3, s.name, { anchor: 'end', size: LABEL_SIZE }));
	});
	return out.join('');
}
