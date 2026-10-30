/**
 * chart-3d-projection.ts: resolve a chart's `c:view3D` into the projection
 * PowerPoint draws it with (an oblique projection for right-angle axes, a
 * perspective camera otherwise). The implementation moved to `ooxml-core/chart`
 * so Word, Excel and Visio share it; this module re-exports it for the WebGL
 * scene and its tests.
 *
 * @module chart-3d-projection
 */
export {
	resolveChart3DProjection,
	type Chart3DObliqueProjection,
	type Chart3DPerspectiveProjection,
	type Chart3DProjection,
} from 'ooxml-core/chart';
