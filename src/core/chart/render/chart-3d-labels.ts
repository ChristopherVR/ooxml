// Data labels on 3-D cartesian charts (bar3D, line3D, area3D): projects each mark's 3-D anchor
// into screen space through the active camera and delegates text rendering to `chart-svg-labels`.
import type { Chart3DCamera } from './chart-3d-camera';
import type { Chart3DLabelAnchor } from './chart-3d-marks';
import { draw, measure, pointLabels, type PointMark } from './chart-svg-labels';
import type { ChartViewModel } from './chart-view';

const GAP = 3;

/** Data labels of 3-D bars, lines and areas, painted on top of all 3-D meshes. */
export function chart3DDataLabelsSvg(
	model: ChartViewModel,
	camera: Chart3DCamera,
	anchors: readonly Chart3DLabelAnchor[],
): string {
	if (!anchors.length) return '';
	const out: string[] = [];
	const points: PointMark[] = [];
	const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
	for (const anchor of anchors) {
		if (anchor.kind === 'point') {
			const q = camera.project(anchor.point);
			points.push({
				series: anchor.series,
				index: anchor.index,
				x: q.x,
				y: q.y,
				radius: 2,
			});
			continue;
		}
		const label = anchor.series.labels?.[anchor.index];
		if (!label) continue;
		const m = measure(model, label);
		let position = label.position ?? (stacked ? 'ctr' : 'outEnd');
		if (stacked && position === 'outEnd') position = 'ctr';
		const cMid = (anchor.c0 + anchor.c1) / 2;
		const zMid = (anchor.z0 + anchor.z1) / 2;
		if (model.horizontal) {
			const dir = anchor.negative ? -1 : 1;
			if (position === 'ctr') {
				const p = camera.project([(anchor.v0 + anchor.v1) / 2, cMid, anchor.z0]);
				draw(m, label, p.x, p.y, 'middle', out);
			} else if (position === 'inBase') {
				const p = camera.project([anchor.v0, cMid, anchor.z0]);
				draw(m, label, p.x + dir * GAP, p.y, dir > 0 ? 'start' : 'end', out);
			} else if (position === 'inEnd') {
				const p = camera.project([anchor.v1, cMid, anchor.z0]);
				draw(m, label, p.x - dir * GAP, p.y, dir > 0 ? 'end' : 'start', out);
			} else {
				const p = camera.project([anchor.v1, cMid, zMid]);
				draw(m, label, p.x + dir * GAP, p.y, dir > 0 ? 'start' : 'end', out);
			}
			continue;
		}
		const dir = anchor.negative ? 1 : -1;
		if (position === 'ctr') {
			const p = camera.project([cMid, (anchor.v0 + anchor.v1) / 2, anchor.z0]);
			draw(m, label, p.x, p.y, 'middle', out);
		} else if (position === 'inBase') {
			const p = camera.project([cMid, anchor.v0, anchor.z0]);
			draw(m, label, p.x, p.y + dir * (GAP + m.h / 2), 'middle', out);
		} else if (position === 'inEnd') {
			const p = camera.project([cMid, anchor.v1, anchor.z0]);
			draw(m, label, p.x, p.y - dir * (GAP + m.h / 2), 'middle', out);
		} else {
			const p = camera.project([cMid, anchor.v1, zMid]);
			draw(m, label, p.x, p.y + dir * (GAP + m.h / 2), 'middle', out);
		}
	}
	if (points.length) pointLabels(model, points, out);
	return out.join('');
}
