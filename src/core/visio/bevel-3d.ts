/**
 * Bevel and 3-D Rotation (Home > Shape Styles > Effects). Visio keeps them in shape cells:
 * BevelTopType, BevelTopWidth and BevelTopHeight (Bevel section of the ShapeSheet) and
 * RotationXAngle, RotationYAngle, RotationZAngle and Perspective (3-D Rotation).
 * The cell names and the 1-12 bevel numbering were read from Visio 16 through COM and its raster
 * export: type 1 is the gallery's first bevel (Circle), 2 Relaxed Inset, 3 Cross, 4 Cool Slant,
 * 5 Angle, 6 Soft Round, 7 Convex, 8 Slope, 9 Divot, 10 Riblet, 11 Hard Edge, 12 Art Deco.
 * Visio's own SVG export ignores both effects; its raster export draws them.
 */
import { fail } from './package-common';
import type { FormattingWrite } from './edit-formatting';
import { number, type Cells } from './sheet';

/** A top bevel: Visio type 1-12 and its width and height in inches. */
export interface VisioBevel {
	type: number;
	width: number;
	height: number;
}
/** 3-D rotation in radians (counter-clockwise Z); `perspective` is the camera field of view. */
export interface VisioRotation3d {
	x: number;
	y: number;
	z: number;
	perspective: number;
	keepTextFlat: boolean;
}

export const VISIO_BEVEL_PRESETS = [
	['circle', 'Circle'],
	['relaxed-inset', 'Relaxed Inset'],
	['cross', 'Cross'],
	['cool-slant', 'Cool Slant'],
	['angle', 'Angle'],
	['soft-round', 'Soft Round'],
	['convex', 'Convex'],
	['slope', 'Slope'],
	['divot', 'Divot'],
	['riblet', 'Riblet'],
	['hard-edge', 'Hard Edge'],
	['art-deco', 'Art Deco'],
].map(([id, label], index) => ({ id: id!, label: label!, type: index + 1 }));
/** The gallery's default bevel size, in points. */
export const VISIO_BEVEL_POINTS = 6;
/** Bevel types drawn with the light inverted (an inset) and with a hard, narrow profile. */
export const VISIO_BEVEL_TYPES_INSET: readonly number[] = [2, 9];
export const VISIO_BEVEL_TYPES_SHARP: readonly number[] = [3, 4, 5, 8, 10, 11, 12];

/** Presets are named by `none` or one of the ids above. */
export type VisioBevelPreset = string;
export const isVisioBevelPreset = (value: unknown): value is VisioBevelPreset =>
	value === 'none' || VISIO_BEVEL_PRESETS.some((preset) => preset.id === value);

/** Rotation presets in degrees. Positive X turns the right side away and positive Y the top. */
export const VISIO_ROTATION_PRESETS: readonly {
	id: string;
	label: string;
	x: number;
	y: number;
	z: number;
	perspective: number;
}[] = [
	['off', 'No Rotation', 0, 0, 0, 0],
	['isometric-left-down', 'Isometric Left Down', -45, -35, 0, 0],
	['isometric-right-up', 'Isometric Right Up', 45, 35, 0, 0],
	['isometric-top-up', 'Isometric Top Up', 0, 45, 0, 0],
	['off-axis-left', 'Off Axis Left', -30, 20, 0, 0],
	['off-axis-right', 'Off Axis Right', 30, 20, 0, 0],
	['perspective-left', 'Perspective Left', -35, 0, 0, 60],
	['perspective-right', 'Perspective Right', 35, 0, 0, 60],
	['perspective-relaxed', 'Perspective Relaxed', 20, 10, 0, 60],
].map(([id, label, x, y, z, perspective]) => ({
	id: id as string,
	label: label as string,
	x: x as number,
	y: y as number,
	z: z as number,
	perspective: perspective as number,
}));
export type VisioRotationPreset = string;
export const isVisioRotationPreset = (value: unknown): value is VisioRotationPreset =>
	VISIO_ROTATION_PRESETS.some((preset) => preset.id === value);

type Add = (
	name: string,
	value: string | number,
	category: FormattingWrite['category'],
	unit?: string,
	formula?: string,
) => void;

/** Cell writes for a bevel preset; lengths are written in points as Visio does. */
export function bevelWrites(id: VisioBevelPreset, add: Add): void {
	const preset = VISIO_BEVEL_PRESETS.find((candidate) => candidate.id === id);
	add('BevelTopType', preset?.type ?? 0, 'FillStyle');
	add('BevelTopWidth', preset ? VISIO_BEVEL_POINTS / 72 : 0, 'FillStyle', 'PT');
	add('BevelTopHeight', preset ? VISIO_BEVEL_POINTS / 72 : 0, 'FillStyle', 'PT');
}
/** Cell writes for a 3-D rotation preset; angles are stored in radians with the DEG unit. */
export function rotationWrites(id: VisioRotationPreset, add: Add): void {
	const preset = VISIO_ROTATION_PRESETS.find((candidate) => candidate.id === id);
	if (!preset) return;
	const radians = (degrees: number) => (degrees * Math.PI) / 180;
	add('RotationXAngle', radians(preset.x), 'FillStyle', 'DEG');
	add('RotationYAngle', radians(preset.y), 'FillStyle', 'DEG');
	add('RotationZAngle', radians(preset.z), 'FillStyle', 'DEG');
	add('Perspective', radians(preset.perspective), 'FillStyle', 'DEG');
}
export function snapshotBevel(value: unknown): VisioBevelPreset {
	if (!isVisioBevelPreset(value)) fail('INVALID_EDIT', 'Unknown bevel preset.');
	return value;
}
export function snapshotRotation(value: unknown): VisioRotationPreset {
	if (!isVisioRotationPreset(value)) fail('INVALID_EDIT', 'Unknown 3-D rotation preset.');
	return value;
}

const saved = (cells: Cells, name: string) =>
	cells.has(name) && cells.get(name)?.value !== 'Themed' ? number(cells, name, 0) : 0;
/** The saved top bevel; a `Themed` cell takes its value from the theme, which is not read. */
export function shapeBevel(cells: Cells): VisioBevel | undefined {
	const type = Math.trunc(saved(cells, 'BevelTopType'));
	const width = Math.max(0, Math.min(2, saved(cells, 'BevelTopWidth')));
	const height = Math.max(0, Math.min(2, saved(cells, 'BevelTopHeight')));
	return type >= 1 && type <= 12 && (width > 0 || height > 0) ? { type, width, height } : undefined;
}
/** Whether a bevel the model does not draw is saved: a bottom bevel or a contour. */
export const hasUndrawnBevel = (cells: Cells): boolean =>
	saved(cells, 'BevelBottomType') > 0 || saved(cells, 'BevelContourSize') > 0;
export function shapeRotation3d(cells: Cells): VisioRotation3d | undefined {
	const angle = (name: string) => Math.max(-Math.PI, Math.min(Math.PI, saved(cells, name)));
	const x = angle('RotationXAngle'),
		y = angle('RotationYAngle'),
		z = angle('RotationZAngle');
	if (!x && !y && !z) return undefined;
	const flat = cells.get('KeepTextFlat')?.value?.trim().toUpperCase();
	return {
		x,
		y,
		z,
		perspective: Math.max(0, Math.min(Math.PI, saved(cells, 'Perspective'))),
		keepTextFlat: flat === 'TRUE' || flat === '1',
	};
}
/**
 * The parallel projection of a flat shape turned by the rotation, as an SVG matrix over local
 * y-up coordinates [a, b, c, d] (x' = a*x + c*y, y' = b*x + d*y). Measured against Visio 16's
 * raster export: X compresses the width, Y the height, X and Y together shear the shape so its
 * right side drops, and Z turns it counter-clockwise. Perspective foreshortening is not applied.
 */
export function visioRotation3dMatrix(
	rotation: Pick<VisioRotation3d, 'x' | 'y' | 'z'>,
): [number, number, number, number] {
	const [ca, sa] = [Math.cos(rotation.x), Math.sin(rotation.x)];
	const [cb, sb] = [Math.cos(rotation.y), Math.sin(rotation.y)];
	const [cc, sc] = [Math.cos(rotation.z), Math.sin(rotation.z)];
	return [cc * ca + sc * sa * sb, sc * ca - cc * sa * sb, -sc * cb, cc * cb];
}
