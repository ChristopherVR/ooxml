// Mesh generation for the 3-D bar / column, line (ribbon) and area charts, from a chart view model
// and its {@link Chart3DLayout}. Bars are boxes, line segments are thin parallelogram prisms and
// areas are prisms of their outline; series stand in depth rows when the grouping is `standard`.
import { boxMesh, extrudeMesh, type Chart3DMesh } from './chart-3d-mesh';
import type { Chart3DLayout } from './chart-3d-layout';
import { chart3DSeriesInDepth } from './chart-3d-layout';
import { plotted } from './chart-svg-cartesian';
import type { ChartViewModel } from './chart-view';

/** Thickness of a line ribbon as a fraction of the value axis length. */
const RIBBON_THICKNESS = 0.025;

const stackedOf = (model: ChartViewModel): boolean =>
	model.grouping === 'stacked' || model.grouping === 'percentStacked';

/** Bars and columns: boxes clustered, stacked or standing in depth rows. */
export function chart3DBarMeshes(model: ChartViewModel, layout: Chart3DLayout): Chart3DMesh[] {
	const bars = layout.bars;
	if (!bars) return [];
	const values = plotted(model);
	const stacked = stackedOf(model);
	const clustered = model.grouping === 'clustered';
	const inDepth = chart3DSeriesInDepth(model);
	const pos = new Array<number>(model.categories.length).fill(0);
	const neg = new Array<number>(model.categories.length).fill(0);
	const base = layout.value(0);
	const meshes: Chart3DMesh[] = [];
	model.series.forEach((s, si) => {
		const slot = clustered ? si : 0;
		const [z0, z1] = layout.rowZ(inDepth ? si : 0);
		values[si]?.forEach((v, i) => {
			if (v === null || i >= model.categories.length) return;
			let from = 0;
			let to = v;
			if (stacked) {
				const acc = v >= 0 ? pos : neg;
				from = acc[i] ?? 0;
				to = from + v;
				acc[i] = to;
			}
			const a = stacked ? layout.value(from) : base;
			const b = layout.value(to);
			const centre = (i + 0.5) * layout.band;
			const meta = {
				series: s.sourceIndex ?? si,
				point: i,
				color: s.pointColors?.[i] ?? s.color,
			};
			const mesh = layout.horizontal
				? boxMesh(
						a,
						b,
						centre -
							(bars.clusterWidth / 2 - bars.singleBarWidth - slot * bars.step) -
							bars.singleBarWidth,
						centre - (bars.clusterWidth / 2 - bars.singleBarWidth - slot * bars.step),
						z0,
						z1,
						meta,
					)
				: boxMesh(
						centre - bars.clusterWidth / 2 + slot * bars.step,
						centre - bars.clusterWidth / 2 + slot * bars.step + bars.singleBarWidth,
						a,
						b,
						z0,
						z1,
						meta,
					);
			if (mesh) meshes.push(mesh);
		});
	});
	return meshes;
}

/** Lines: one ribbon prism per segment between consecutive present points. */
export function chart3DLineMeshes(model: ChartViewModel, layout: Chart3DLayout): Chart3DMesh[] {
	const values = plotted(model);
	const stacked = stackedOf(model);
	const inDepth = chart3DSeriesInDepth(model);
	const thick = layout.valueLength * RIBBON_THICKNESS;
	const acc = new Array<number>(model.categories.length).fill(0);
	const meshes: Chart3DMesh[] = [];
	model.series.forEach((s, si) => {
		const [z0, z1] = layout.rowZ(inDepth ? si : 0);
		const pts = (values[si] ?? []).map((v, i): [number, number] | null => {
			if (v === null || i >= model.categories.length) return null;
			const top = stacked ? (acc[i] = (acc[i] ?? 0) + v) : v;
			return [(i + 0.5) * layout.band, layout.value(top)];
		});
		for (let i = 0; i + 1 < pts.length; i++) {
			const p = pts[i];
			const q = pts[i + 1];
			if (!p || !q) continue;
			const mesh = extrudeMesh([p, q, [q[0], q[1] - thick], [p[0], p[1] - thick]], z0, z1, {
				series: s.sourceIndex ?? si,
				point: i,
				color: s.color,
			});
			if (mesh) meshes.push(mesh);
		}
	});
	return meshes;
}

/** Areas: a prism of each series' outline down to the baseline (or the stack below it). */
export function chart3DAreaMeshes(model: ChartViewModel, layout: Chart3DLayout): Chart3DMesh[] {
	const values = plotted(model);
	const stacked = stackedOf(model);
	const inDepth = chart3DSeriesInDepth(model);
	const cats = model.categories.length;
	const x = (i: number): number =>
		cats > 1 ? (i / (cats - 1)) * layout.catLength : layout.catLength / 2;
	const base = layout.value(0);
	const acc = new Array<number>(cats).fill(0);
	const meshes: Chart3DMesh[] = [];
	model.series.forEach((s, si) => {
		const [z0, z1] = layout.rowZ(inDepth ? si : 0);
		const upper: [number, number][] = [];
		const lower: [number, number][] = [];
		for (let i = 0; i < cats; i++) {
			const v = values[si]?.[i] ?? 0;
			const below = stacked ? (acc[i] ?? 0) : 0;
			if (stacked) acc[i] = below + v;
			upper.push([x(i), layout.value(below + v)]);
			lower.push([x(i), stacked ? layout.value(below) : base]);
		}
		const mesh = extrudeMesh([...upper, ...lower.reverse()], z0, z1, {
			series: s.sourceIndex ?? si,
			point: 0,
			color: s.color,
		});
		if (mesh) meshes.push(mesh);
	});
	return meshes;
}
