// The 3-D chart painter: dispatches a chart view model that carries a `c:view3D` to the pie, or to
// the cartesian scene (walls, marks, labels) for bars, lines, areas and surfaces.
import { createChart3DCamera } from './chart-3d-camera';
import { chart3DLayout } from './chart-3d-layout';
import { chart3DAreaMeshes, chart3DBarMeshes, chart3DLineMeshes } from './chart-3d-marks';
import type { Chart3DMesh } from './chart-3d-mesh';
import { chart3DPieSvg } from './chart-3d-pie';
import { resolveChart3DProjection } from './chart-3d-projection';
import { chart3DMeshesSvg } from './chart-3d-scene';
import { chart3DSurfaceMeshes } from './chart-3d-surface';
import { chart3DLabelMargins, chart3DLabelsSvg, chart3DWallsSvg } from './chart-3d-walls';
import { chartAreaRect } from './chart-svg-appearance';
import type { Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

const PROJECTION_TYPE: Partial<Record<ChartViewModel['type'], string>> = {
	bar: 'bar3D',
	column: 'bar3D',
	line: 'line3D',
	area: 'area3D',
	pie: 'pie3D',
	surface: 'surface',
};

/** Whether the painter draws this view model as a 3-D chart. */
export function chart3DApplies(model: ChartViewModel): boolean {
	return model.view3D !== undefined && PROJECTION_TYPE[model.type] !== undefined;
}

/** A 3-D chart into `area`. Call only when {@link chart3DApplies}. */
export function chart3DSvg(model: ChartViewModel, area: Rect): string {
	const view = model.view3D ?? {};
	if (model.type === 'pie')
		return chartAreaRect(model, 'plotArea', area) + chart3DPieSvg(model, area, view);
	const projection = resolveChart3DProjection(PROJECTION_TYPE[model.type] ?? 'bar3D', view);
	const margins = chart3DLabelMargins(model);
	const inner: Rect = {
		x: area.x + margins.left,
		y: area.y + 6,
		w: Math.max(1, area.w - margins.left - 8),
		h: Math.max(1, area.h - margins.bottom - 10),
	};
	const layout = chart3DLayout(model, inner, view);
	const camera = createChart3DCamera(projection, layout.size, inner);
	let meshes: Chart3DMesh[];
	if (model.type === 'surface') meshes = chart3DSurfaceMeshes(model, layout);
	else if (model.type === 'line') meshes = chart3DLineMeshes(model, layout);
	else if (model.type === 'area') meshes = chart3DAreaMeshes(model, layout);
	else meshes = chart3DBarMeshes(model, layout);
	const edge = model.type === 'area' || model.type === 'surface';
	const cats = model.categories.length;
	const catPos = edge
		? (i: number) => (cats > 1 ? (i / (cats - 1)) * layout.catLength : layout.catLength / 2)
		: (i: number) => (i + 0.5) * layout.band;
	return [
		chartAreaRect(model, 'plotArea', area),
		chart3DWallsSvg(model, layout, camera),
		chart3DMeshesSvg(camera, meshes),
		chart3DLabelsSvg(model, layout, camera, catPos),
	].join('');
}
