/**
 * The 3-D Rotation gallery (Home > Shape Styles > Effects > 3-D Rotation), DOM-free.
 *
 * Values are the Office preset cameras (ECMA-376 `presetCameraType`) read from PowerPoint's
 * `ThreeD` object over COM and converted to Visio's cell convention. Comparing renders of the same
 * triangle in both products (Visio 16 raster export against PowerPoint's) showed that Visio's
 * `RotationXAngle` and `RotationZAngle` have the opposite sign of PowerPoint's `RotationX` and
 * `RotationZ`, and `RotationYAngle` the same sign. So each entry below is `[-x, y, -z]` of the
 * PowerPoint value, with `Perspective` the preset's field of view (45, 80 degrees for the extreme
 * heroic cameras) or 0 for parallel cameras. Visio's gallery could not be driven headlessly, so
 * the membership and the order follow the Office gallery (Parallel, then Perspective); the oblique
 * cameras are omitted because they carry no rotation (all three angles read 0). The values Visio
 * writes when its own gallery is used were not read.
 */
export interface VisioRotationPresetDefinition {
	id: string;
	label: string;
	/** Degrees, Visio's cell convention. */
	x: number;
	y: number;
	z: number;
	perspective: number;
}

const rows: readonly (readonly [string, string, number, number, number, number])[] = [
	['off', 'No Rotation', 0, 0, 0, 0],
	// Parallel
	['isometric-left-down', 'Isometric Left Down', 45, 35, 0, 0],
	['isometric-right-up', 'Isometric Right Up', -45, 35, 0, 0],
	['isometric-top-up', 'Isometric Top Up', -45.28, -35.4, 60.16, 0],
	['isometric-bottom-down', 'Isometric Bottom Down', -45.28, 35.4, -60.16, 0],
	['off-axis-1-left', 'Off Axis 1 Left', 64, 18, 0, 0],
	['off-axis-1-right', 'Off Axis 1 Right', -26, 18, 0, 0],
	['off-axis-1-top', 'Off Axis 1 Top', -53.45, -58.74, 57.64, 0],
	['off-axis-2-left', 'Off Axis 2 Left', 26, 18, 0, 0],
	['off-axis-2-right', 'Off Axis 2 Right', -64, 18, 0, 0],
	['off-axis-2-top', 'Off Axis 2 Top', 53.45, -58.74, -57.64, 0],
	// Perspective
	['perspective-above', 'Perspective Above', 0, -20, 0, 45],
	['perspective-below', 'Perspective Below', 0, 20, 0, 45],
	['perspective-left', 'Perspective Left', 20, 0, 0, 45],
	['perspective-right', 'Perspective Right', -20, 0, 0, 45],
	['perspective-relaxed', 'Perspective Relaxed', 0, -50.44, 0, 45],
	['perspective-relaxed-moderately', 'Perspective Relaxed Moderately', 0, -35.16, 0, 45],
	['perspective-contrasting-left', 'Perspective Contrasting Left', 43.94, 10.4, -3.55, 45],
	['perspective-contrasting-right', 'Perspective Contrasting Right', -43.94, 10.4, 3.55, 45],
	['perspective-heroic-left', 'Perspective Heroic Left', 14.29, -11.02, 2.63, 45],
	['perspective-heroic-right', 'Perspective Heroic Right', -14.29, -11.02, -2.63, 45],
	['perspective-heroic-extreme-left', 'Perspective Heroic Extreme Left', 34.46, 8.12, -2.91, 80],
	['perspective-heroic-extreme-right', 'Perspective Heroic Extreme Right', -34.46, 8.12, 2.91, 80],
];

export const VISIO_ROTATION_PRESETS: readonly VisioRotationPresetDefinition[] = rows.map(
	([id, label, x, y, z, perspective]) => ({ id, label, x, y, z, perspective }),
);
