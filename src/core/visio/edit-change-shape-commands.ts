import { fail } from './package-common';
import { isVisioBasicShape, type VisioBasicShape } from './basic-shapes';

/** A Basic Shapes outline, or the native Ellipse row (`circle` keeps the shape's own size). */
export type VisioChangeShapeTarget = VisioBasicShape | 'ellipse' | 'circle';

/**
 * Visio's Home > Editing > Change Shape for one local 2D shape: its Geometry sections are
 * replaced by the target outline. Position, size, rotation, flips, text and formatting stay.
 */
export interface VisioChangeShapeEdit {
	type: 'change-shape';
	pageId: string;
	shapeId: string;
	shape: VisioChangeShapeTarget;
}

export const isVisioChangeShapeTarget = (value: unknown): value is VisioChangeShapeTarget =>
	value === 'ellipse' || value === 'circle' || isVisioBasicShape(value);

/** Copy a change-shape command, keeping only its known fields. */
export function snapshotChangeShape(edit: VisioChangeShapeEdit): VisioChangeShapeEdit {
	if (
		typeof edit.shapeId !== 'string' ||
		!/^[1-9]\d{0,9}$/.test(edit.shapeId) ||
		Number(edit.shapeId) > 4294967295
	)
		fail('INVALID_EDIT', 'Change Shape needs a canonical positive unsigned shape ID.');
	if (!isVisioChangeShapeTarget(edit.shape)) fail('INVALID_EDIT', 'Unknown basic shape.');
	return { type: edit.type, pageId: edit.pageId, shapeId: edit.shapeId, shape: edit.shape };
}
