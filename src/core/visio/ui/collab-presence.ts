// What a Visio collaborator broadcasts through the shared collaboration session: the page they
// are on, their pointer in drawing coordinates and the shapes they have selected. DOM-free; the
// viewer passes `sanitizeVisioPresence` to `openCollaborationSession` so every peer validates
// what it receives the same way, and a host-owned session should be created with it too.
import type { VisioMatrix, VisioPage, VisioShape } from '../model';

/** The presence payload of one Visio collaborator. */
export interface VisioPresence {
	/** Page id of the page being viewed; an empty string while no page is open. */
	page: string;
	/** Pointer in the page's SVG drawing units (inches, y down); absent while outside the page. */
	x?: number | undefined;
	y?: number | undefined;
	/** Shape ids selected on that page (bounded). */
	selection?: string[] | undefined;
}

export const VISIO_PRESENCE_MAX_SELECTION = 64;
export const VISIO_PRESENCE_MAX_ID = 128;
/** Drawing coordinates beyond this are corrupt, not a drawing (a Visio page is at most ~600 in). */
export const VISIO_PRESENCE_MAX_COORDINATE = 100_000;

const clean = (value: unknown): string | undefined =>
	typeof value === 'string' && value.length > 0 ? value.slice(0, VISIO_PRESENCE_MAX_ID) : undefined;

const coordinate = (value: unknown): number | undefined =>
	typeof value === 'number' && Number.isFinite(value)
		? Math.max(-VISIO_PRESENCE_MAX_COORDINATE, Math.min(VISIO_PRESENCE_MAX_COORDINATE, value))
		: undefined;

/**
 * Validate a raw awareness `presence` entry. Returns `null` (hide the peer's location) when it
 * names no usable page; wrong-typed pointer or selection fields are dropped, lists are bounded.
 */
export function sanitizeVisioPresence(raw: Record<string, unknown>): VisioPresence | null {
	const page =
		typeof raw.page === 'number' && Number.isFinite(raw.page) ? String(raw.page) : raw.page;
	if (typeof page !== 'string') return null;
	const presence: VisioPresence = { page: page.slice(0, VISIO_PRESENCE_MAX_ID) };
	const x = coordinate(raw.x);
	const y = coordinate(raw.y);
	if (x !== undefined && y !== undefined) {
		presence.x = x;
		presence.y = y;
	}
	if (Array.isArray(raw.selection)) {
		const ids = new Set<string>();
		for (const entry of raw.selection) {
			const id = clean(entry);
			if (id) ids.add(id);
			if (ids.size >= VISIO_PRESENCE_MAX_SELECTION) break;
		}
		if (ids.size) presence.selection = [...ids];
	}
	return presence;
}

/** The presence a window starts with, before it has a page or a pointer. */
export const VISIO_INITIAL_PRESENCE: VisioPresence = Object.freeze({ page: '' });

/** A shape's frame in page drawing space: `transform` maps its local y-up box to the page. */
export interface VisioShapeFrame {
	transform: VisioMatrix;
	width: number;
	height: number;
}

function multiply(p: VisioMatrix, c: VisioMatrix): VisioMatrix {
	return [
		p[0] * c[0] + p[2] * c[1],
		p[1] * c[0] + p[3] * c[1],
		p[0] * c[2] + p[2] * c[3],
		p[1] * c[2] + p[3] * c[3],
		p[0] * c[4] + p[2] * c[5] + p[4],
		p[1] * c[4] + p[3] * c[5] + p[5],
	];
}

/** The frame of shape `id` on `page` (parents composed in), or undefined when it is not there. */
export function visioShapeFrame(page: VisioPage, id: string): VisioShapeFrame | undefined {
	const identity: VisioMatrix = [1, 0, 0, 1, 0, 0];
	const find = (
		shapes: readonly VisioShape[],
		parent: VisioMatrix,
	): VisioShapeFrame | undefined => {
		for (const shape of shapes) {
			const transform = multiply(parent, shape.transform);
			if (shape.id === id) return { transform, width: shape.width, height: shape.height };
			const inner = shape.children.length ? find(shape.children, transform) : undefined;
			if (inner) return inner;
		}
		return undefined;
	};
	const frame = find(page.shapes, identity);
	return frame && [...frame.transform, frame.width, frame.height].every(Number.isFinite)
		? frame
		: undefined;
}
