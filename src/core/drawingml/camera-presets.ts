/**
 * The DrawingML camera presets (`a:camera/@prst`) as approximate rotate X/Y/Z angles and a
 * reference perspective distance, DOM-free so every product's 3D painter shares one table.
 * Moved from the PowerPoint renderer (`render/visual-3d-camera-presets`, which re-exports it).
 *
 * There, these angles are the rendered transform only for a preset with no measured homography,
 * or when an explicit `a:camera/a:rot`, `@fov` or `@zoom` override is present; for other presets
 * they are an extrusion panel visibility and shading direction hint. The shared SmartArt painter
 * uses them as the pose of a preset that has no fitted diagram camera (`diagram/scene-camera`).
 */

/**
 * Camera preset configuration: reference perspective distance (px, at
 * `visual-3d-camera-fov`'s `REFERENCE_SIZE_PX`) and base rotation angles (in degrees). These
 * approximate the OOXML camera preset positions.
 */
export interface CameraPresetConfig {
	/** Reference CSS perspective distance in px at `visual-3d-camera-fov`'s `REFERENCE_SIZE_PX`, or `undefined` for a parallel (non-perspective) projection. */
	perspectiveRefPx?: number;
	rotateX: number;
	rotateY: number;
	rotateZ: number;
	/** COM-calibrated `perspective-origin` (% of the element's box); only used on the explicit-override/no-homography fallback path. */
	perspectiveOriginXPct?: number;
	perspectiveOriginYPct?: number;
}

/**
 * Reference perspective distance (px, at `visual-3d-camera-fov`'s `REFERENCE_SIZE_PX`) used for
 * an explicit rotation with no camera preset at all. Matches the pre-existing
 * flat `800px` fallback.
 */
export const DEFAULT_CUSTOM_REF_PX = 800;

export const CAMERA_PRESET_MAP: Record<string, CameraPresetConfig> = {
	orthographicFront: { rotateX: 0, rotateY: 0, rotateZ: 0 },
	perspectiveFront: { perspectiveRefPx: 1000, rotateX: 0, rotateY: 0, rotateZ: 0 },
	perspectiveAbove: { perspectiveRefPx: 1000, rotateX: 20, rotateY: 0, rotateZ: 0 },
	perspectiveBelow: { perspectiveRefPx: 1000, rotateX: -20, rotateY: 0, rotateZ: 0 },
	perspectiveLeft: { perspectiveRefPx: 1000, rotateX: 0, rotateY: -20, rotateZ: 0 },
	perspectiveRight: { perspectiveRefPx: 1000, rotateX: 0, rotateY: 20, rotateZ: 0 },
	perspectiveAboveLeftFacing: { perspectiveRefPx: 1000, rotateX: 20, rotateY: -25, rotateZ: 0 },
	perspectiveAboveRightFacing: { perspectiveRefPx: 1000, rotateX: 20, rotateY: 25, rotateZ: 0 },
	perspectiveContrastingLeftFacing: {
		perspectiveRefPx: 800,
		rotateX: 15,
		rotateY: -30,
		rotateZ: 0,
		perspectiveOriginXPct: 93,
		perspectiveOriginYPct: 0,
	},
	perspectiveContrastingRightFacing: {
		perspectiveRefPx: 800,
		rotateX: 15,
		rotateY: 30,
		rotateZ: 0,
		perspectiveOriginXPct: 7,
		perspectiveOriginYPct: 0,
	},
	perspectiveHeroicLeftFacing: { perspectiveRefPx: 600, rotateX: 10, rotateY: -35, rotateZ: 0 },
	perspectiveHeroicRightFacing: { perspectiveRefPx: 600, rotateX: 10, rotateY: 35, rotateZ: 0 },
	perspectiveHeroicExtremeLeftFacing: {
		perspectiveRefPx: 500,
		rotateX: 8,
		rotateY: -45,
		rotateZ: 0,
		perspectiveOriginXPct: 32,
		perspectiveOriginYPct: 42,
	},
	perspectiveHeroicExtremeRightFacing: {
		perspectiveRefPx: 500,
		rotateX: 8,
		rotateY: 45,
		rotateZ: 0,
		perspectiveOriginXPct: 68,
		perspectiveOriginYPct: 42,
	},
	perspectiveRelaxed: { perspectiveRefPx: 1200, rotateX: 10, rotateY: 0, rotateZ: 0 },
	perspectiveRelaxedModerately: { perspectiveRefPx: 1400, rotateX: 5, rotateY: 0, rotateZ: 0 },
	// No `perspectiveRefPx` on any isometric* preset: COM-measured real
	// PowerPoint output shows a true parallelogram (parallel projection, not
	// perspective) for these, see `visual-3d-camera`'s module doc comment.
	isometricLeftDown: { rotateX: -35, rotateY: 45, rotateZ: 0 },
	isometricRightUp: { rotateX: -35, rotateY: -45, rotateZ: 0 },
	isometricLeftUp: { rotateX: 35, rotateY: 45, rotateZ: 0 },
	isometricRightDown: { rotateX: 35, rotateY: -45, rotateZ: 0 },
	isometricTopUp: { rotateX: -55, rotateY: 0, rotateZ: 45 },
	isometricTopDown: { rotateX: -55, rotateY: 0, rotateZ: -45 },
	isometricBottomUp: { rotateX: 55, rotateY: 0, rotateZ: 45 },
	isometricBottomDown: { rotateX: 55, rotateY: 0, rotateZ: -45 },
	isometricOffAxis1Left: { rotateX: -30, rotateY: 30, rotateZ: 0 },
	isometricOffAxis1Right: { rotateX: -30, rotateY: -30, rotateZ: 0 },
	isometricOffAxis1Top: { rotateX: -45, rotateY: 0, rotateZ: 30 },
	isometricOffAxis2Left: { rotateX: -30, rotateY: 20, rotateZ: 0 },
	isometricOffAxis2Right: { rotateX: -30, rotateY: -20, rotateZ: 0 },
	isometricOffAxis2Top: { rotateX: -45, rotateY: 0, rotateZ: -30 },
	isometricOffAxis3Left: { rotateX: -25, rotateY: 35, rotateZ: 0 },
	isometricOffAxis3Right: { rotateX: -25, rotateY: -35, rotateZ: 0 },
	isometricOffAxis3Bottom: { rotateX: 45, rotateY: 0, rotateZ: 30 },
	isometricOffAxis4Left: { rotateX: -25, rotateY: 25, rotateZ: 0 },
	isometricOffAxis4Right: { rotateX: -25, rotateY: -25, rotateZ: 0 },
	isometricOffAxis4Bottom: { rotateX: 45, rotateY: 0, rotateZ: -30 },
	obliqueTopLeft: { perspectiveRefPx: 900, rotateX: -20, rotateY: 20, rotateZ: 0 },
	obliqueTop: { perspectiveRefPx: 900, rotateX: -25, rotateY: 0, rotateZ: 0 },
	obliqueTopRight: { perspectiveRefPx: 900, rotateX: -20, rotateY: -20, rotateZ: 0 },
	obliqueLeft: { perspectiveRefPx: 900, rotateX: 0, rotateY: 25, rotateZ: 0 },
	obliqueRight: { perspectiveRefPx: 900, rotateX: 0, rotateY: -25, rotateZ: 0 },
	obliqueBottomLeft: { perspectiveRefPx: 900, rotateX: 20, rotateY: 20, rotateZ: 0 },
	obliqueBottom: { perspectiveRefPx: 900, rotateX: 25, rotateY: 0, rotateZ: 0 },
	obliqueBottomRight: { perspectiveRefPx: 900, rotateX: 20, rotateY: -20, rotateZ: 0 },
	// Pre-2007 (legacy) WordArt/AutoShape 3-D camera names. Schema-legal but not
	// exposed in the modern "3-D Rotation" gallery; they only surface via files
	// round-tripped from legacy Office or hand-authored XML.
	legacyObliqueTopLeft: { perspectiveRefPx: 900, rotateX: -20, rotateY: 20, rotateZ: 0 },
	legacyObliqueTop: { perspectiveRefPx: 900, rotateX: -25, rotateY: 0, rotateZ: 0 },
	legacyObliqueTopRight: { perspectiveRefPx: 900, rotateX: -20, rotateY: -20, rotateZ: 0 },
	legacyObliqueLeft: { perspectiveRefPx: 900, rotateX: 0, rotateY: 25, rotateZ: 0 },
	legacyObliqueFront: { perspectiveRefPx: 900, rotateX: 0, rotateY: 0, rotateZ: 0 },
	legacyObliqueRight: { perspectiveRefPx: 900, rotateX: 0, rotateY: -25, rotateZ: 0 },
	legacyObliqueBottomLeft: { perspectiveRefPx: 900, rotateX: 20, rotateY: 20, rotateZ: 0 },
	legacyObliqueBottom: { perspectiveRefPx: 900, rotateX: 25, rotateY: 0, rotateZ: 0 },
	legacyObliqueBottomRight: { perspectiveRefPx: 900, rotateX: 20, rotateY: -20, rotateZ: 0 },
	legacyPerspectiveTopLeft: { perspectiveRefPx: 700, rotateX: -30, rotateY: 30, rotateZ: 0 },
	legacyPerspectiveTop: { perspectiveRefPx: 700, rotateX: -35, rotateY: 0, rotateZ: 0 },
	legacyPerspectiveTopRight: { perspectiveRefPx: 700, rotateX: -30, rotateY: -30, rotateZ: 0 },
	legacyPerspectiveLeft: { perspectiveRefPx: 700, rotateX: 0, rotateY: 35, rotateZ: 0 },
	legacyPerspectiveFront: { perspectiveRefPx: 700, rotateX: 0, rotateY: 0, rotateZ: 0 },
	legacyPerspectiveRight: { perspectiveRefPx: 700, rotateX: 0, rotateY: -35, rotateZ: 0 },
	legacyPerspectiveBottomLeft: { perspectiveRefPx: 700, rotateX: 30, rotateY: 30, rotateZ: 0 },
	legacyPerspectiveBottom: { perspectiveRefPx: 700, rotateX: 35, rotateY: 0, rotateZ: 0 },
	legacyPerspectiveBottomRight: { perspectiveRefPx: 700, rotateX: 30, rotateY: -30, rotateZ: 0 },
};
