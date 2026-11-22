import { fail } from './package-common';
import type { VisioPathSegment } from './path-fit';
export type { VisioPathSegment } from './path-fit';

/**
 * Create a 2D shape from a drawn path (Freeform, Pencil and Arc tools). Drawing inches with a
 * bottom-left origin; (`x`, `y`) is the start point. A closed path must end on its start point.
 */
export interface VisioPathStroke {
	x: number;
	y: number;
	segments: VisioPathSegment[];
}

export interface VisioPathCreateEdit extends VisioPathStroke {
	type: 'create-path';
	pageId: string;
	shapeId: string;
	closed?: boolean;
	subpaths?: VisioPathStroke[];
	stroke?: { name: string; color: string; weight: number };
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
	let segmentCount = 0;
	const snapshotSegments = (segments: VisioPathSegment[]): VisioPathSegment[] => {
		if (
			!Array.isArray(segments) ||
			!segments.length ||
			(segmentCount += segments.length) > VISIO_PATH_SEGMENT_LIMIT
		)
			invalid(`A path needs between 1 and ${VISIO_PATH_SEGMENT_LIMIT} segments.`);
		return Array.from(segments, snapshotSegment);
	};
	if (edit.closed !== undefined && typeof edit.closed !== 'boolean')
		invalid('Path closure must be a boolean.');
	const snapshotSegment = (segment: VisioPathSegment): VisioPathSegment => {
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
	};
	const segments = snapshotSegments(edit.segments);
	if (
		edit.subpaths !== undefined &&
		(!Array.isArray(edit.subpaths) ||
			edit.subpaths.length > VISIO_PATH_SEGMENT_LIMIT ||
			edit.closed)
	)
		invalid('A compound path must have bounded, unfilled strokes.');
	const subpaths = edit.subpaths?.map((path) => {
		if (!path || typeof path !== 'object') invalid('Invalid path stroke.');
		return { x: value(path.x), y: value(path.y), segments: snapshotSegments(path.segments) };
	});
	const stroke = edit.stroke;
	if (
		stroke &&
		(typeof stroke.name !== 'string' ||
			!stroke.name.trim() ||
			stroke.name.length > 128 ||
			/[\u0000-\u001f]/u.test(stroke.name) ||
			typeof stroke.color !== 'string' ||
			!/^#[0-9a-f]{6}$/i.test(stroke.color) ||
			!(value(stroke.weight) > 0))
	)
		invalid('Invalid path stroke style.');
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
		...(subpaths?.length ? { subpaths } : {}),
		...(stroke
			? { stroke: { name: stroke.name, color: stroke.color, weight: stroke.weight } }
			: {}),
		...(edit.closed === undefined ? {} : { closed: edit.closed }),
	};
}
