// Paints 3-D meshes as SVG with the painter's algorithm: marks far to near, and inside a mark the
// faces turned away from the viewer are culled and the rest drawn far to near with the flat
// per-face shading Office uses.
import type { Chart3DCamera } from './chart-3d-camera';
import type { Chart3DMesh, Chart3DMeshFace, Vec3 } from './chart-3d-mesh';
import { shadeChart3DFace } from './chart-3d-shading';
import { n } from './chart-svg-util';

const centroid = (pts: readonly Vec3[]): Vec3 => {
	const sum = pts.reduce<[number, number, number]>(
		(acc, p) => [acc[0] + p[0], acc[1] + p[1], acc[2] + p[2]],
		[0, 0, 0],
	);
	return [sum[0] / pts.length, sum[1] / pts.length, sum[2] / pts.length];
};

/** Whether a face turns toward the viewer. Double-sided faces always do. */
export function chart3DFaceVisible(camera: Chart3DCamera, face: Chart3DMeshFace): boolean {
	if (!face.normal) return true;
	const d = camera.viewDir(centroid(face.pts));
	return face.normal[0] * d[0] + face.normal[1] * d[1] + face.normal[2] * d[2] < 0;
}

/** A projected polygon with the paint the 3-D meshes use (a hairline of the fill hides seams). */
export function chart3DPolygon(
	camera: Chart3DCamera,
	pts: readonly Vec3[],
	fill: string,
	extra = '',
): string {
	const points = pts
		.map((p) => {
			const q = camera.project(p);
			return `${n(q.x)},${n(q.y)}`;
		})
		.join(' ');
	return `<polygon points="${points}" fill="${fill}" stroke="${fill}" stroke-width="0.5" stroke-linejoin="round"${extra}/>`;
}

/** The meshes far to near, each a `<g>` carrying its series and point for hit testing. */
export function chart3DMeshesSvg(camera: Chart3DCamera, meshes: readonly Chart3DMesh[]): string {
	const out: string[] = [];
	const ordered = [...meshes].sort((a, b) => camera.order(a.center, b.center));
	for (const mesh of ordered) {
		const faces = mesh.faces
			.filter((face) => chart3DFaceVisible(camera, face))
			.map((face) => ({ face, depth: camera.project(centroid(face.pts)).depth }))
			.sort((a, b) => b.depth - a.depth);
		if (!faces.length) continue;
		out.push(`<g data-chart-series="${mesh.series}" data-chart-point="${mesh.point}">`);
		for (const { face } of faces)
			out.push(
				chart3DPolygon(camera, face.pts, face.fill ?? shadeChart3DFace(mesh.color, face.kind)),
			);
		out.push('</g>');
	}
	return out.join('');
}
