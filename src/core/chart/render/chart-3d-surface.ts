// Mesh generation for a 3-D surface chart: categories along x, series along z, the cached values
// as heights. Every grid cell is a double-sided patch coloured by the value band it falls in and
// lit by its slope, as Office's banded surface is.
import type { Chart3DMesh, Vec3 } from './chart-3d-mesh';
import type { Chart3DLayout } from './chart-3d-layout';
import { shade } from '../style-palette';
import type { ChartViewModel } from './chart-view';

const LIGHT: Vec3 = (() => {
	const v: Vec3 = [-0.3, 0.8, -0.5];
	const len = Math.hypot(...v);
	return [v[0] / len, v[1] / len, v[2] / len];
})();
const HEX = /^#[0-9a-f]{6}$/i;
/** Band colours when the series carry none a surface legend could use. */
const BAND_FALLBACK = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];

export function chart3DSurfaceMeshes(model: ChartViewModel, layout: Chart3DLayout): Chart3DMesh[] {
	const cols = model.categories.length;
	const rows = model.series.length;
	const scale = model.valueAxis;
	if (cols < 2 || rows < 1 || !scale) return [];
	const zRows = rows === 1 ? 2 : rows;
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
			const band = Math.max(
				0,
				Math.min(
					Math.max(0, scale.ticks.length - 2),
					Math.floor((avg / layout.valueLength) * Math.max(1, scale.ticks.length - 1)),
				),
			);
			const color =
				model.series[band % rows]?.color && HEX.test(model.series[band % rows]!.color)
					? model.series[band % rows]!.color
					: (BAND_FALLBACK[band % BAND_FALLBACK.length] ?? '#4472C4');
			const [a, b, , d] = pts as [Vec3, Vec3, Vec3, Vec3];
			const e1: Vec3 = [b[0] - a[0], b[1] - a[1], b[2] - a[2]];
			const e2: Vec3 = [d[0] - a[0], d[1] - a[1], d[2] - a[2]];
			const nx = e1[1] * e2[2] - e1[2] * e2[1];
			const ny = e1[2] * e2[0] - e1[0] * e2[2];
			const nz = e1[0] * e2[1] - e1[1] * e2[0];
			const len = Math.hypot(nx, ny, nz) || 1;
			const lit = Math.abs((nx * LIGHT[0] + ny * LIGHT[1] + nz * LIGHT[2]) / len);
			meshes.push({
				series: Math.min(r, rows - 1),
				point: c,
				color,
				center: [(a[0] + b[0]) / 2, avg, (a[2] + d[2]) / 2],
				faces: [{ pts, normal: undefined, kind: 'top', fill: shade(color, 0.35 * (1 - lit)) }],
			});
		}
	return meshes;
}
