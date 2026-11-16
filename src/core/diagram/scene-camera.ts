/**
 * The whole-diagram camera of a SmartArt scene quick style (`dgm:styleDef/dgm:scene3d`), DOM-free,
 * shared by the PowerPoint 3D scene and the Word and Excel SVG painter. Moved from the PowerPoint
 * renderer (`render/smartart-3d-scene-camera`, which re-exports it).
 *
 * Rotation convention (degrees): a diagram point `P` is turned by `R = Rz(rev) . Rx(lat) .
 * Ry(-lon)` about the diagram centre, then projected along -z. Its orthographic front-face
 * projection is the `a:camera/a:rot` model measured on PowerPoint's own exports (`x = X cos lon`,
 * `y = Y cos lat - X sin lat sin lon`, then a `rev` roll), so an explicit `a:rot` is used as
 * authored.
 *
 * Presets without `a:rot` take their pose from {@link DIAGRAM_SCENE_CAMERA_PRESETS}: unit-square
 * homographies measured on PowerPoint's rendered output (COM exports), fitted to this pinhole model
 * (rotation and camera distance, in multiples of the measured 2in square's side). The distance is
 * absolute, so it is converted to layout px ({@link MEASURED_SQUARE_PX}) rather than scaled to the
 * diagram. A preset missing from the table falls back to the shared `CAMERA_PRESET_MAP` angles.
 * `@fov` is not modelled.
 */
import { CAMERA_PRESET_MAP } from '../drawingml/camera-presets';
import type { DrawingScene3D } from '../drawingml/text-3d-model';

/** A resolved diagram camera; `distance` and `zoom` in layout px and as a fraction. */
export interface DiagramSceneCamera {
	projection: 'orthographic' | 'perspective';
	latDeg: number;
	lonDeg: number;
	revDeg: number;
	/** Camera distance from the diagram centre in layout px (`perspective` only). */
	distance: number;
	/** `a:camera/@zoom` as a fraction (1 = 100%). */
	zoom: number;
}

/** A fitted preset pose; `distance` absent means a parallel projection. */
export interface DiagramScenePresetPose {
	lat: number;
	lon: number;
	rev: number;
	distance?: number;
}

/** Camera distance used for a perspective preset with no fitted pose. */
const DEFAULT_PERSPECTIVE_DISTANCE = 6;

/** Side of the measurement square the preset distances are in: 2in at 96 px/in. */
export const MEASURED_SQUARE_PX = 192;

/** Presets the SmartArt scene quick styles use (plus their mirror images). */
export const DIAGRAM_SCENE_CAMERA_PRESETS: Readonly<Record<string, DiagramScenePresetPose>> = {
	orthographicFront: { lat: 0, lon: 0, rev: 0 },
	perspectiveFront: { lat: 0, lon: 0, rev: 0, distance: 160 },
	perspectiveLeft: { lat: 0, lon: 20.4, rev: 0, distance: 8.3 },
	perspectiveRight: { lat: 0, lon: -20.4, rev: 0, distance: 8.3 },
	perspectiveRelaxed: { lat: -50.8, lon: 0.3, rev: -0.2, distance: 7.8 },
	perspectiveRelaxedModerately: { lat: -35.1, lon: 0, rev: 0, distance: 7.9 },
	perspectiveHeroicExtremeRightFacing: { lat: 8, lon: -36.1, rev: 2.1, distance: 4.2 },
	perspectiveHeroicExtremeLeftFacing: { lat: 8, lon: 36.1, rev: -2.1, distance: 4.2 },
	isometricOffAxis2Left: { lat: 18, lon: 26, rev: 0 },
	isometricOffAxis2Right: { lat: 18, lon: -26, rev: 0 },
};

const DEG_PER_UNIT = 60000;
const toDeg = (value: number | undefined): number => (value ? value / DEG_PER_UNIT : 0);

/** `a:camera/@zoom` as a fraction; a raw ST_PositivePercentage (`95000`) is converted. */
function normalizeZoom(value: number | undefined): number {
	if (!value || value <= 0) return 1;
	return value > 10 ? value / 100000 : value;
}

function presetPose(preset: string): DiagramScenePresetPose {
	const fitted = DIAGRAM_SCENE_CAMERA_PRESETS[preset];
	if (fitted) return fitted;
	const legacy = CAMERA_PRESET_MAP[preset];
	return {
		lat: legacy?.rotateX ?? 0,
		lon: legacy?.rotateY ?? 0,
		rev: legacy?.rotateZ ?? 0,
		...(preset.startsWith('perspective') ? { distance: DEFAULT_PERSPECTIVE_DISTANCE } : {}),
	};
}

/**
 * The diagram camera of a scene, or `undefined` when it leaves the diagram face-on (no scene, or
 * an `orthographicFront` camera with no rotation), where flat framing already matches Office.
 */
export function resolveDiagramSceneCamera(
	scene: DrawingScene3D | undefined,
): DiagramSceneCamera | undefined {
	if (!scene) return undefined;
	const pose = presetPose(scene.cameraPreset ?? 'orthographicFront');
	const hasRot =
		scene.cameraRotX !== undefined ||
		scene.cameraRotY !== undefined ||
		scene.cameraRotZ !== undefined;
	const latDeg = hasRot ? toDeg(scene.cameraRotX) : pose.lat;
	const lonDeg = hasRot ? toDeg(scene.cameraRotY) : pose.lon;
	const revDeg = hasRot ? toDeg(scene.cameraRotZ) : pose.rev;
	const distance = pose.distance === undefined ? undefined : pose.distance * MEASURED_SQUARE_PX;
	const zoom = normalizeZoom(scene.cameraZoom);
	if (latDeg === 0 && lonDeg === 0 && revDeg === 0 && distance === undefined && zoom === 1)
		return undefined;
	return {
		projection: distance === undefined ? 'orthographic' : 'perspective',
		latDeg,
		lonDeg,
		revDeg,
		distance: distance ?? 0,
		zoom,
	};
}

/** The camera's rotation as a row-major 3x3 matrix (`R = Rz(rev) . Rx(lat) . Ry(-lon)`). */
export function diagramSceneCameraMatrix(camera: DiagramSceneCamera): number[] {
	const rad = Math.PI / 180;
	const cl = Math.cos(camera.latDeg * rad);
	const sl = Math.sin(camera.latDeg * rad);
	const co = Math.cos(-camera.lonDeg * rad);
	const so = Math.sin(-camera.lonDeg * rad);
	const cr = Math.cos(camera.revDeg * rad);
	const sr = Math.sin(camera.revDeg * rad);
	// Rx(lat) . Ry(-lon)
	const m = [co, 0, so, sl * so, cl, -sl * co, -cl * so, sl, cl * co] as const;
	// Rz(rev) . m
	return [
		cr * m[0] - sr * m[3],
		cr * m[1] - sr * m[4],
		cr * m[2] - sr * m[5],
		sr * m[0] + cr * m[3],
		sr * m[1] + cr * m[4],
		sr * m[2] + cr * m[5],
		m[6],
		m[7],
		m[8],
	];
}

/**
 * Project a diagram point (layout px, `z` towards the viewer) through the camera about `centre`:
 * rotate, then a parallel or pinhole projection, then the zoom. Returns screen px.
 */
export function projectDiagramScenePoint(
	camera: DiagramSceneCamera,
	centre: { x: number; y: number },
	point: { x: number; y: number; z: number },
): { x: number; y: number } {
	const m = diagramSceneCameraMatrix(camera);
	// The rotation model is y-up; screen y grows downwards.
	const x = point.x - centre.x;
	const y = centre.y - point.y;
	const X = m[0]! * x + m[1]! * y + m[2]! * point.z;
	const Y = m[3]! * x + m[4]! * y + m[5]! * point.z;
	const Z = m[6]! * x + m[7]! * y + m[8]! * point.z;
	const scale =
		camera.projection === 'perspective' && camera.distance > 0
			? camera.distance / Math.max(camera.distance - Z, camera.distance * 0.05)
			: 1;
	return {
		x: centre.x + X * scale * camera.zoom,
		y: centre.y - Y * scale * camera.zoom,
	};
}
