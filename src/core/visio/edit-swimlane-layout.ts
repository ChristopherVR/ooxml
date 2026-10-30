import { attribute, children } from './sheet';
import { fail } from './package-common';
import { cells, numeric } from './edit-geometry-cells';
import { applyGeometryEdit } from './edit-geometry';
import { boundsOf, placeVisioBox, type Box } from './edit-container-fit';
import {
	containerClosure,
	containerMembersOf,
	shapeIdOf,
	topLevelShapes,
	userValue,
} from './edit-container-members';
import { VISIO_SWIMLANE_FRAME_ROW, VISIO_SWIMLANE_LANE_ROW } from './shape-structure';
import type { VisioSwimlaneDirection } from './edit-swimlane-commands';
import { stackSwimlanes } from './swimlane-stack';
import { VISIO_SEPARATOR_ROW, setUserRow } from './edit-swimlane-shapes';
export interface Scope {
	roots: ReadonlyMap<string, Element>;
	document: Element;
	root: Element;
	pageId: string;
	check: () => void;
	/** Whether the lanes carry their shapes (local ones) or the caller moves them. */
	contents: 'move' | 'leave';
}
export interface Options {
	direction: VisioSwimlaneDirection;
	title: boolean;
	separators: boolean;
	rotate: boolean;
	style: number;
}

const flag = (shape: Element, name: string, fallback: boolean): boolean => {
	const value = userValue(shape, name);
	return value === undefined ? fallback : value !== '0';
};
export function frameOf(scope: Scope, frameId: string): Element {
	const frame = topLevelShapes(scope.root).find((shape) => shapeIdOf(shape) === frameId);
	if (!frame || userValue(frame, VISIO_SWIMLANE_FRAME_ROW) !== '1')
		fail('EDIT_TARGET_NOT_FOUND', 'A cross-functional flowchart frame is required.');
	return frame;
}
export function optionsOf(frame: Element): Options {
	const style = Number(userValue(frame, 'visCFFStyle'));
	return {
		direction: flag(frame, 'CFFVertical', false) ? 'vertical' : 'horizontal',
		title: flag(frame, 'visShowTitle', true),
		separators: flag(frame, 'visShowPhase', true),
		rotate: flag(frame, 'visRotateLabel', false),
		style: Number.isInteger(style) && style >= 1 && style <= 4 ? style : 1,
	};
}
export const isLane = (shape: Element) => userValue(shape, VISIO_SWIMLANE_LANE_ROW) === '1';
export const isSeparator = (shape: Element) => userValue(shape, VISIO_SEPARATOR_ROW) === '1';

/** The frame's lanes in stacking order: top to bottom, or left to right. */
export function lanesOf(
	scope: Scope,
	frame: Element,
	direction: VisioSwimlaneDirection,
): Element[] {
	const members = new Set(containerMembersOf(frame));
	const lanes = topLevelShapes(scope.root).filter(
		(shape) => members.has(shapeIdOf(shape)) && isLane(shape),
	);
	const key = (lane: Element) => {
		const box = boundsOf([lane]);
		return box ? (direction === 'horizontal' ? -box.maxY : box.minX) : 0;
	};
	return lanes.sort((a, b) => key(a) - key(b));
}
export function laneOf(
	scope: Scope,
	frame: Element,
	laneId: string,
	direction: VisioSwimlaneDirection,
) {
	const lane = lanesOf(scope, frame, direction).find((item) => shapeIdOf(item) === laneId);
	if (!lane) fail('EDIT_TARGET_NOT_FOUND', 'The shape is not a lane of this flowchart.');
	return lane;
}

/** Page-space connectors with a glued end follow the shapes they are glued to. */
function gluedConnectors(root: Element): Set<string> {
	const ids = new Set<string>();
	for (const connects of children(root, 'Connects'))
		for (const connect of children(connects, 'Connect')) {
			const id = attribute(connect, 'FromSheet');
			if (id) ids.add(id);
		}
	return ids;
}
function shift(scope: Scope, ids: Iterable<string>, dx: number, dy: number): void {
	if (Math.abs(dx) < 1e-9 && Math.abs(dy) < 1e-9) return;
	const glued = gluedConnectors(scope.root);
	const byId = new Map(topLevelShapes(scope.root).map((shape) => [shapeIdOf(shape), shape]));
	for (const id of ids) {
		const shape = byId.get(id);
		if (!shape || glued.has(id)) continue;
		const local = cells(shape);
		applyGeometryEdit(
			scope.roots,
			scope.document,
			{
				type: 'move-shape',
				pageId: scope.pageId,
				shapeId: id,
				x: numeric(local.get('PinX')) + dx,
				y: numeric(local.get('PinY')) + dy,
			},
			scope.check,
		);
	}
}

/** Give a lane, separator or frame its box; a sheet the geometry editor refuses stops the edit. */
export function place(scope: Scope, shape: Element, target: Box): void {
	if (!placeVisioBox(scope.roots, scope.document, scope.pageId, shape, target, scope.check))
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'A flowchart shape could not be resized.');
}

/**
 * Stack the lanes edge to edge under the frame's title (the frame's top-left corner stays put),
 * carry each lane's shapes with it unless the caller moves them (contents: 'leave'), resize the
 * phase separators to span the lanes and fit the frame. 'order' overrides the geometric order
 * (a lane just inserted).
 */
export function relayout(scope: Scope, frame: Element, order?: readonly Element[]): void {
	const options = optionsOf(frame);
	const horizontal = options.direction === 'horizontal';
	const frameBox = boundsOf([frame]);
	if (!frameBox) fail('UNSUPPORTED_GEOMETRY_EDIT', 'The frame has no usable bounds.');
	const lanes = (order ?? lanesOf(scope, frame, options.direction)).filter(
		(lane) => boundsOf([lane]) !== undefined,
	);
	const stacking = stackSwimlanes(
		options.direction,
		frameBox,
		options.title,
		lanes.map((lane) => ({ id: shapeIdOf(lane), box: boundsOf([lane])! })),
	);
	const carried = new Map(
		stacking.lanes.map((entry, index) => [index, [...containerClosure(scope.root, lanes[index]!)]]),
	);
	if (scope.contents === 'move') for (const ids of carried.values()) assertMovable(scope.root, ids);
	stacking.lanes.forEach((entry, index) => {
		scope.check();
		place(scope, lanes[index]!, entry.target);
		if (scope.contents === 'move') shift(scope, carried.get(index)!, entry.dx, entry.dy);
	});
	for (const shape of topLevelShapes(scope.root)) {
		if (!isSeparator(shape) || !containerMembersOf(frame).includes(shapeIdOf(shape))) continue;
		const bar = boundsOf([shape]);
		if (!bar) continue;
		place(
			scope,
			shape,
			horizontal
				? { ...bar, minY: stacking.end, maxY: stacking.top }
				: { ...bar, minX: frameBox.minX, maxX: stacking.frame.maxX },
		);
	}
	place(scope, frame, stacking.frame);
	setUserRow(frame, 'numLanes', String(lanes.length), 'STR');
}

/** Stencil instances move through the asynchronous instance path, so the caller must move them. */
function assertMovable(root: Element, ids: readonly string[]): void {
	const byId = new Map(topLevelShapes(root).map((shape) => [shapeIdOf(shape), shape]));
	for (const id of ids)
		if (byId.get(id)?.hasAttribute('Master'))
			fail(
				'UNSUPPORTED_GEOMETRY_EDIT',
				'Stencil shapes in a lane must be moved with move-shape edits (contents: leave).',
			);
}

/** Frame, lanes and separators sit behind the page's other shapes, in that order. */
export function sendBehind(root: Element, shapes: readonly Element[]): void {
	const container = children(root, 'Shapes')[0]!;
	const first = children(container, 'Shape').find((shape) => !shapes.includes(shape));
	for (const shape of shapes) container.insertBefore(shape, first ?? null);
}
