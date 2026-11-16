/**
 * Whole-diagram camera for the SmartArt scene quick styles (framework-agnostic, pure). The camera
 * model, the fitted preset poses and the rotation matrix are the shared, DOM-free
 * `ooxml-core/diagram` scene camera (`diagram/scene-camera.ts`, which documents the convention and
 * the measurements), which Word and Excel's SVG painter uses too; this module keeps the
 * PowerPoint renderer's names and adds the eye position its lit scene needs.
 *
 * @module render/smartart-3d-scene-camera
 */
import {
	DIAGRAM_SCENE_CAMERA_PRESETS,
	MEASURED_SQUARE_PX,
	diagramSceneCameraMatrix,
	resolveDiagramSceneCamera,
} from 'ooxml-core/diagram';

import type { SmartArt3DCamera } from './smartart-3d-solid-types';
import type { Vec3 } from './smartart-3d-types';

export { MEASURED_SQUARE_PX };

/** Presets the SmartArt scene quick styles use (plus their mirror images). */
export const SCENE_CAMERA_PRESETS = DIAGRAM_SCENE_CAMERA_PRESETS;

/**
 * The diagram camera for a quick style's `dgm:scene3d`, or `undefined` when
 * it leaves the diagram face-on.
 */
export const resolveSmartArt3DCamera: (
	scene: Parameters<typeof resolveDiagramSceneCamera>[0],
) => SmartArt3DCamera | undefined = resolveDiagramSceneCamera;

/** The camera's rotation as a row-major 3x3 matrix (`R = Rz(rev) . Rx(lat) . Ry(-lon)`). */
export const smartArt3DCameraMatrix: (camera: SmartArt3DCamera) => number[] =
	diagramSceneCameraMatrix;

/**
 * The camera position in diagram space (before the rotation), for per-point
 * view directions: `R^T . (0, 0, distance)`. `undefined` for a parallel
 * projection.
 */
export function smartArt3DEyeInDiagram(camera: SmartArt3DCamera): Vec3 | undefined {
	if (camera.projection !== 'perspective') {
		return undefined;
	}
	const m = smartArt3DCameraMatrix(camera);
	// Third row of R, i.e. the third column of R^T, times the distance.
	return { x: m[6] * camera.distance, y: m[7] * camera.distance, z: m[8] * camera.distance };
}
