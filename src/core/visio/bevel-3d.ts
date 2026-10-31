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
import { VISIO_ROTATION_PRESETS } from './rotation-presets';
import type { ThemeSolid } from './theme-solid';

/** A top bevel: Visio type 1-12 and its width and height in inches. */
export interface VisioBevel {
	type: number;
	width: number;
	height: number;
	/** `BevelLightingType` 1-15, when the shape saves a non-default one (see `bevel-shading.ts`). */
	lighting?: number;
	/** `BevelMaterialType` 1-12, when saved and not the default. */
	material?: number;
}
/** Extrusion (Format Shape > 3-D Format > Depth): `BevelDepthSize` in inches and its colour. */
export interface VisioExtrusion {
	depth: number;
	color: string;
}
/** Contour (3-D Format > Contour): an outline of `size` inches drawn around the 3-D shape. */
export interface VisioContour {
	size: number;
	color: string;
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

export { VISIO_ROTATION_PRESETS };
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
/** The depth and contour of the 3-D format: lengths in points, colours as #rrggbb. */
export interface SolidPatch {
	extrusionDepth?: number;
	extrusionColor?: string;
	contourSize?: number;
	contourColor?: string;
}
const rgb = (color: string) =>
	`RGB(${[1, 3, 5].map((index) => parseInt(color.slice(index, index + 2), 16)).join(',')})`;
export function solidWrites(edit: SolidPatch, add: Add): void {
	if (edit.extrusionDepth !== undefined)
		add('BevelDepthSize', edit.extrusionDepth / 72, 'FillStyle', 'PT');
	if (edit.extrusionColor !== undefined)
		add('BevelDepthColor', edit.extrusionColor, 'FillStyle', undefined, rgb(edit.extrusionColor));
	if (edit.contourSize !== undefined)
		add('BevelContourSize', edit.contourSize / 72, 'FillStyle', 'PT');
	if (edit.contourColor !== undefined)
		add('BevelContourColor', edit.contourColor, 'FillStyle', undefined, rgb(edit.contourColor));
}
export function snapshotSolid(edit: SolidPatch): SolidPatch {
	const length = (value: unknown, high: number, what: string) => {
		if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > high)
			fail('INVALID_EDIT', `${what} must be 0 to ${high} points.`);
		return Math.round(value * 100) / 100;
	};
	const color = (value: unknown, what: string) => {
		if (typeof value !== 'string' || !/^#[0-9a-f]{6}$/i.test(value))
			fail('INVALID_EDIT', `${what} requires six hexadecimal digits.`);
		return value.toLowerCase();
	};
	const result: SolidPatch = {};
	if (edit.extrusionDepth !== undefined)
		result.extrusionDepth = length(edit.extrusionDepth, 288, 'Extrusion depth');
	if (edit.extrusionColor !== undefined)
		result.extrusionColor = color(edit.extrusionColor, 'The extrusion colour');
	if (edit.contourSize !== undefined)
		result.contourSize = length(edit.contourSize, 72, 'Contour size');
	if (edit.contourColor !== undefined)
		result.contourColor = color(edit.contourColor, 'The contour colour');
	return result;
}
export function snapshotBevel(value: unknown): VisioBevelPreset {
	if (!isVisioBevelPreset(value)) fail('INVALID_EDIT', 'Unknown bevel preset.');
	return value;
}
export function snapshotRotation(value: unknown): VisioRotationPreset {
	if (!isVisioRotationPreset(value)) fail('INVALID_EDIT', 'Unknown 3-D rotation preset.');
	return value;
}

const themedCell = (cells: Cells, name: string) => cells.get(name)?.value === 'Themed';
const saved = (cells: Cells, name: string) =>
	cells.has(name) && !themedCell(cells, name) ? number(cells, name, 0) : 0;
const clamp = (value: number, high: number) => Math.max(0, Math.min(high, value));
/**
 * The top bevel: the saved cells, or for a `Themed` cell the theme's `a:bevelT`. A bevel of type
 * 1-12 with no size still makes the shape a solid (extrusion and contour need it, measured in
 * Visio 16), so it is returned with zero width and height.
 */
export function shapeBevel(cells: Cells, themed?: ThemeSolid): VisioBevel | undefined {
	const read = (name: string, themeValue: number | undefined) =>
		themedCell(cells, name) ? (themeValue ?? 0) : saved(cells, name);
	const type = Math.trunc(read('BevelTopType', themed?.bevel?.type));
	const width = clamp(read('BevelTopWidth', themed?.bevel?.width), 2);
	const height = clamp(read('BevelTopHeight', themed?.bevel?.height), 2);
	if (type < 1 || type > 12) return undefined;
	const lighting = Math.trunc(saved(cells, 'BevelLightingType'));
	const material = Math.trunc(saved(cells, 'BevelMaterialType'));
	return {
		type,
		width,
		height,
		...(lighting >= 1 && lighting <= 15 ? { lighting } : {}),
		...(material >= 1 && material <= 12 ? { material } : {}),
	};
}
/** The extrusion, only on a shape with a top bevel (a depth alone is not drawn by Visio 16). */
export function shapeExtrusion(
	cells: Cells,
	solid: boolean,
	colorOf: (name: string, fallback: string) => string,
): VisioExtrusion | undefined {
	const depth = clamp(saved(cells, 'BevelDepthSize'), 4);
	return solid && depth > 0 ? { depth, color: colorOf('BevelDepthColor', '#ffffff') } : undefined;
}
export function shapeContour(
	cells: Cells,
	solid: boolean,
	themed: ThemeSolid | undefined,
	colorOf: (name: string, fallback: string) => string,
): VisioContour | undefined {
	const size = clamp(
		themedCell(cells, 'BevelContourSize')
			? (themed?.contour?.size ?? 0)
			: saved(cells, 'BevelContourSize'),
		1,
	);
	const color = themedCell(cells, 'BevelContourColor')
		? (themed?.contour?.color ?? colorOf('BevelContourColor', '#000000'))
		: colorOf('BevelContourColor', '#000000');
	return solid && size > 0 ? { size, color } : undefined;
}
/** Whether a bevel the model does not draw is saved: a bottom bevel (no visible effect in front). */
export const hasUndrawnBevel = (cells: Cells): boolean => saved(cells, 'BevelBottomType') > 0;
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
 * right side drops, and Z turns it counter-clockwise. Perspective and extrusion are `solid-3d.ts`.
 */
export function visioRotation3dMatrix(
	rotation: Pick<VisioRotation3d, 'x' | 'y' | 'z'>,
): [number, number, number, number] {
	const [ca, sa] = [Math.cos(rotation.x), Math.sin(rotation.x)];
	const [cb, sb] = [Math.cos(rotation.y), Math.sin(rotation.y)];
	const [cc, sc] = [Math.cos(rotation.z), Math.sin(rotation.z)];
	return [cc * ca + sc * sa * sb, sc * ca - cc * sa * sb, -sc * cb, cc * cb];
}
