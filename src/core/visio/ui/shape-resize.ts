import type { VisioPage, VisioShape, VisioMatrix } from '../model';
import type { VisioEdit } from '../edit-commands';
import { visioAnchoredResizeGeometry, type VisioResizeAnchor } from '../resize-anchor';
import { visioMovementShape, visioPageDragDelta, type VisioPagePoint } from './shape-move';
import { visioPageEditToDrawing } from './page-edit';

/** A flat local group whose member axes admit independent width/height scaling.
 * Source admission additionally proves resize modes, locks, formulas and geometry. */
function resizableLocalGroup(page: VisioPage, shape: VisioShape): boolean {
	if (
		shape.kind !== 'group' ||
		shape.masterId ||
		!shape.children.length ||
		shape.children.length > 500
	)
		return false;
	const ids = new Set([shape.id]);
	for (const member of shape.children) {
		if (
			ids.has(member.id) ||
			member.kind !== 'shape' ||
			member.masterId ||
			member.children.length ||
			member.layerIds?.length ||
			member.image ||
			member.foreignVector ||
			!member.rotation ||
			Math.abs(member.transform[1]) > 1e-10 ||
			Math.abs(member.transform[2]) > 1e-10 ||
			!visioResizeFrame(member, member, { x: 0, y: 0 })
		)
			return false;
		ids.add(member.id);
	}
	return !page.connectors.some(
		(connection) => ids.has(connection.fromShapeId) || ids.has(connection.toShapeId),
	);
}

export const VISIO_RESIZE_HANDLES = [
	{ id: 'sw', x: 0, y: 0 },
	{ id: 's', x: 0.5, y: 0 },
	{ id: 'se', x: 1, y: 0 },
	{ id: 'e', x: 1, y: 0.5 },
	{ id: 'ne', x: 1, y: 1 },
	{ id: 'n', x: 0.5, y: 1 },
	{ id: 'nw', x: 0, y: 1 },
	{ id: 'w', x: 0, y: 0.5 },
] as const;
export type VisioResizeHandle = (typeof VISIO_RESIZE_HANDLES)[number]['id'];
export interface VisioResizeFrame {
	width: number;
	height: number;
	transform: VisioMatrix;
	pinX: number;
	pinY: number;
}
/** Only orthonormal ordinary local shapes have a scene-proven rectangular resize frame. */
export function visioResizeShape(page: VisioPage, id: string): VisioShape | undefined {
	const shape = visioMovementShape(page, id);
	// Flat local groups use the core member-scaling proof. A stencil group instead follows
	// the formulas in its master; the two source paths keep their own admission rules.
	// Layered local shapes move, but core resize admission still refuses their layer membership;
	// stencil shapes are resized on their (unlocked) stencil layer.
	return shape &&
		(!shape.children.length || !!shape.masterId || resizableLocalGroup(page, shape)) &&
		(!shape.layerIds?.length || !!shape.masterId) &&
		visioResizeFrame(shape, { width: shape.width, height: shape.height }, { x: 0, y: 0 })
		? shape
		: undefined;
}
export function visioResizeFrame(
	shape: VisioShape,
	size: { width: number; height: number },
	anchor: VisioResizeAnchor,
): VisioResizeFrame | undefined {
	if (!shape.rotation) return undefined;
	const geometry = visioAnchoredResizeGeometry(
		{ ...shape, pinX: shape.rotation.pinX, pinY: shape.rotation.pinY },
		size,
		anchor,
	);
	return geometry
		? {
				width: size.width,
				height: size.height,
				transform: geometry.transform,
				pinX: geometry.pinX,
				pinY: geometry.pinY,
			}
		: undefined;
}
/** How a drag maps onto the shape's size. Visio keeps the proportions when a corner handle is
 * dragged, and on any handle while Shift is held (`proportional`); side handles otherwise change
 * one dimension. */
export interface VisioResizeDragOptions {
	proportional?: boolean;
	/** Ctrl: resize about the centre, so each dragged edge moves twice the pointer distance. */
	symmetric?: boolean;
}
/** True when a drag on `handle` keeps the shape's proportions, as Visio does. */
export function visioResizeKeepsProportions(
	handle: VisioResizeHandle,
	options: VisioResizeDragOptions = {},
): boolean {
	return !!options.proportional || handle.length === 2;
}
/** Scale a raw resize uniformly. A corner follows the axis that changed more; a side handle
 * scales the other axis with it about its centre. The ratio survives the minimum-size clamp. */
export function visioProportionalSize(
	shape: { width: number; height: number },
	raw: { width: number; height: number },
	location: { x: number; y: number },
	min = { width: Math.min(1 / 16, shape.width), height: Math.min(1 / 16, shape.height) },
): { width: number; height: number } {
	const sx = raw.width / shape.width,
		sy = raw.height / shape.height;
	let scale =
		location.x === 0.5
			? sy
			: location.y === 0.5
				? sx
				: Math.abs(sx - 1) >= Math.abs(sy - 1)
					? sx
					: sy;
	scale = Math.max(scale, min.width / shape.width, min.height / shape.height);
	return { width: shape.width * scale, height: shape.height * scale };
}
/** Drag in physical page coordinates along the saved rotated/reflected local axes.
 * Crossing the opposite handle clamps at 1/16 physical inch; it does not silently change flips.
 */
export function visioResizeDrag(
	page: VisioPage,
	id: string,
	handle: VisioResizeHandle,
	start: VisioPagePoint,
	end: VisioPagePoint,
	options: VisioResizeDragOptions = {},
): { frame: VisioResizeFrame; command: Extract<VisioEdit, { type: 'resize-shape' }> } | undefined {
	const shape = visioResizeShape(page, id),
		location = VISIO_RESIZE_HANDLES.find((item) => item.id === handle);
	if (!shape || !location) return undefined;
	const delta = visioPageDragDelta(start, end),
		[a, b, c, d] = shape.transform,
		determinant = a * d - b * c;
	const x = (d * delta.x - c * delta.y) / determinant,
		y = (-b * delta.x + a * delta.y) / determinant;
	// A symmetric (Ctrl) drag keeps the centre fixed, so each moving edge travels twice as far.
	const k = options.symmetric ? 2 : 1;
	const raw = {
		width:
			location.x === 0.5
				? shape.width
				: Math.max(Math.min(1 / 16, shape.width), shape.width + k * (location.x === 1 ? x : -x)),
		height:
			location.y === 0.5
				? shape.height
				: Math.max(Math.min(1 / 16, shape.height), shape.height + k * (location.y === 1 ? y : -y)),
	};
	const size = visioResizeKeepsProportions(handle, options)
		? visioProportionalSize(shape, raw, location)
		: raw;
	// The opposite point stays fixed; a proportional side drag grows the other axis about its
	// centre. A symmetric drag anchors the centre instead.
	const anchor: VisioResizeAnchor = options.symmetric
		? { x: 0.5, y: 0.5 }
		: {
				x: (1 - location.x) as VisioResizeAnchor['x'],
				y: (1 - location.y) as VisioResizeAnchor['y'],
			};
	const frame = visioResizeFrame(shape, size, anchor);
	if (!frame) return undefined;
	try {
		const command = visioPageEditToDrawing(page, {
			type: 'resize-shape',
			pageId: page.id,
			shapeId: id,
			...size,
			anchor,
		});
		return command.type === 'resize-shape' && command.width <= 1e6 && command.height <= 1e6
			? { frame, command }
			: undefined;
	} catch {
		return undefined;
	}
}
