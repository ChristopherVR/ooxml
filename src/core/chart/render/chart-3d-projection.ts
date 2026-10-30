// Resolves a chart's `c:view3D` into the projection Office draws it with. Extracted from the
// PowerPoint renderer (`src/ui/src/pptx/render/chart-3d-projection.ts`, see PROVENANCE.md) so
// Word, Excel, Visio and PowerPoint share one answer.
//
// Office's 3-D charts use one of two projections, chosen by `c:view3D/@rAngAx`:
// - right-angle axes (the bar default): an oblique parallel projection. The front plane is drawn
//   flat and each mark's depth recedes along a fixed screen direction set by `rotX` / `rotY`.
// - otherwise (line, area, pie and surface by default): a perspective camera, elevation `rotX`,
//   azimuth `rotY`, field of view from `c:perspective`.
// An absent `rAngAx` defaults per type: true for bar3D, false for the others. pie3D defaults to a
// 30/0 tilt, every other type to 15/20.

/** The `c:view3D` fields the projection reads (a PowerPoint or neutral chart model both fit). */
export interface Chart3DViewInput {
	rotX?: number | undefined;
	rotY?: number | undefined;
	/** `c:perspective`, 0..240 degrees. */
	perspective?: number | undefined;
	/** `c:rAngAx`: right-angle axes. */
	rAngAx?: boolean | undefined;
	/** `c:depthPercent`: depth as a percentage of the base width (20..2000, default 100). */
	depthPercent?: number | undefined;
	/** `c:hPercent`: height as a percentage of the base width (5..500). */
	hPercent?: number | undefined;
	/** `c:gapDepth` of the chart group (0..500, default 150). */
	gapDepth?: number | undefined;
}

/** Right-angle axes: flat front faces, marks recede along a fixed screen-space shear. */
export interface Chart3DObliqueProjection {
	mode: 'oblique';
	/** World-X screen shift per world unit of depth. */
	shearX: number;
	/** World-Y screen shift per world unit of depth (negative when seen from above). */
	shearY: number;
	rotXDeg: number;
	rotYDeg: number;
}

/** A perspective camera framing the whole 3-D box. */
export interface Chart3DPerspectiveProjection {
	mode: 'perspective';
	rotXDeg: number;
	rotYDeg: number;
	/** Resolved vertical field of view in degrees. */
	fovDeg: number;
}

export type Chart3DProjection = Chart3DObliqueProjection | Chart3DPerspectiveProjection;

const DEFAULT_ROT_X = 15;
const DEFAULT_ROT_Y = 20;
const DEFAULT_PIE_ROT_X = 30;
const DEFAULT_PIE_ROT_Y = 0;
const DEFAULT_PERSPECTIVE_DEG = 30;
/** `c:perspective` is documented 0-240; Office's own picker stops at 120. */
const PERSPECTIVE_MAX_DEG = 120;
const FOV_MIN = 15;
const FOV_MAX = 75;

const clamp = (value: number, min: number, max: number): number =>
	Math.min(Math.max(value, min), max);

function resolvePerspectiveFov(perspectiveDeg: number | undefined): number {
	const perspective = clamp(perspectiveDeg ?? DEFAULT_PERSPECTIVE_DEG, 0, PERSPECTIVE_MAX_DEG);
	return FOV_MIN + (perspective / PERSPECTIVE_MAX_DEG) * (FOV_MAX - FOV_MIN);
}

/** Resolve the projection of a 3-D chart from its type (`bar3D`, `pie3D`...) and `c:view3D`. */
export function resolveChart3DProjection(
	chartType: string,
	view3D: Chart3DViewInput | undefined,
): Chart3DProjection {
	const isPie = chartType === 'pie3D';
	const rotXDeg = view3D?.rotX ?? (isPie ? DEFAULT_PIE_ROT_X : DEFAULT_ROT_X);
	const rotYDeg = view3D?.rotY ?? (isPie ? DEFAULT_PIE_ROT_Y : DEFAULT_ROT_Y);
	const rAngAx = view3D?.rAngAx ?? chartType === 'bar3D';
	if (rAngAx) {
		const rx = (rotXDeg * Math.PI) / 180;
		const ry = (rotYDeg * Math.PI) / 180;
		return { mode: 'oblique', shearX: Math.sin(ry), shearY: -Math.sin(rx), rotXDeg, rotYDeg };
	}
	return {
		mode: 'perspective',
		rotXDeg,
		rotYDeg,
		fovDeg: resolvePerspectiveFov(view3D?.perspective),
	};
}
