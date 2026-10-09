import { fail } from './package-common';
import type { VisioPathSegment } from './path-fit';
export type { VisioPathSegment } from './path-fit';

/**
 * Create a 2D shape from a drawn path (Freeform, Pencil and Arc tools). Drawing inches with a
 * bottom-left origin; (`x`, `y`) is the start point. A closed path must end on its start point.
 */
export interface VisioPathCreateEdit {
	type: 'create-path';
	pageId: string;
	shapeId: string;
	x: number;
	y: number;
	segments: VisioPathSegment[];
	closed?: boolean;
}

export const VISIO_PATH_SEGMENT_LIMIT = 2000;

/** Bound, clone and validate a path creation; strips host properties. */
export function snapshotPathCreation(
	edit: VisioPathCreateEdit,
	invalid: (message: string) => never = (message) => fail('INVALID_EDIT', message),
): VisioPathCreateEdit {
	const value = (input: unknown): number => {
		if (typeof input !== 'number' || !Number.isFinite(input) || Math.abs(input) > 1e6)
			invalid('Path coordinates must be finite drawing inches within limits.');
		return input;
	};
	if (
		!Array.isArray(edit.segments) ||
		!edit.segments.length ||
		edit.segments.length > VISIO_PATH_SEGMENT_LIMIT
	)
		invalid(`A path needs between 1 and ${VISIO_PATH_SEGMENT_LIMIT} segments.`);
	if (edit.closed !== undefined && typeof edit.closed !== 'boolean')
		invalid('Path closure must be a boolean.');
	const segments = Array.from(edit.segments, (segment): VisioPathSegment => {
		if (!segment || typeof segment !== 'object') invalid('Invalid path segment.');
		const end = { x: value(segment.x), y: value(segment.y) };
		switch (segment.kind) {
			case 'line':
				return { kind: 'line', ...end };
			case 'cubic':
				return {
					kind: 'cubic',
					...end,
					x1: value(segment.x1),
					y1: value(segment.y1),
					x2: value(segment.x2),
					y2: value(segment.y2),
				};
			case 'arc': {
				const ratio = value(segment.ratio);
				if (!(ratio > 0)) invalid('An arc needs a positive axis ratio.');
				return { kind: 'arc', ...end, a: value(segment.a), b: value(segment.b), ratio };
			}
			default:
				return invalid('Unknown path segment kind.');
		}
	});
	const x = value(edit.x),
		y = value(edit.y);
	const last = segments[segments.length - 1]!;
	if (edit.closed && Math.hypot(last.x - x, last.y - y) > 1e-9)
		invalid('A closed path must end on its start point.');
	return {
		type: 'create-path',
		pageId: edit.pageId,
		shapeId: edit.shapeId,
		x,
		y,
		segments,
		...(edit.closed === undefined ? {} : { closed: edit.closed }),
	};
}
