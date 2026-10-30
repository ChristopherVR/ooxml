import { attribute, children } from './sheet';
import { fail } from './package-common';
import { cells, numeric } from './edit-geometry-cells';
import { applyGeometryEdit } from './edit-geometry';
import { boundsOf, placeVisioBox, type Box } from './edit-container-fit';
import {
	containerClosure,
	containerMembersOf,
	listMembersOf,
	setRelation,
	shapeIdOf,
	topLevelShapes,
	userValue,
	writeContainerMembers,
	writeListMembers,
} from './edit-container-members';
import {
	VISIO_SEPARATOR_ROW,
	VISIO_SWIMLANE_FRAME_ROW,
	VISIO_SWIMLANE_LANE_ROW,
} from './shape-structure';
import { VISIO_RELATIONSHIP_LIST_SIBLING } from './container-relationships';
import type { VisioSwimlaneDirection } from './edit-swimlane-commands';
import { stackSwimlanes } from './swimlane-stack';
import { createListShape, setUserRow } from './edit-swimlane-shapes';
import { materializeCells, recalculateMasterInstances } from './cff-recalc';
import type { CffMasterSet, CffRole } from './cff-masters';
export interface Scope {
	roots: ReadonlyMap<string, Element>;
	document: Element;
	root: Element;
	pageId: string;
	check: () => void;
	/** Whether the lanes carry their shapes (local ones) or the caller moves them. */
	contents: 'move' | 'leave';
	/** Visio's cross-functional flowchart masters in the drawing, when it has them. */
	masters: CffMasterSet;
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
/** The role a shape has by the master it is an instance of (a chart Visio built). */
export const masterRole = (scope: Scope, shape: Element): CffRole | undefined =>
	scope.masters.roleOf(attribute(shape, 'Master'));
export function frameOf(scope: Scope, frameId: string): Element {
	const frame = topLevelShapes(scope.root).find((shape) => shapeIdOf(shape) === frameId);
	if (
		!frame ||
		(userValue(frame, VISIO_SWIMLANE_FRAME_ROW) !== '1' && masterRole(scope, frame) !== 'frame')
	)
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
const hasCategory = (shape: Element, category: string) =>
	(userValue(shape, 'msvShapeCategories') ?? '').split(';').includes(category);
export const isLane = (scope: Scope, shape: Element) =>
	userValue(shape, VISIO_SWIMLANE_LANE_ROW) === '1' ||
	hasCategory(shape, 'Swimlane') ||
	masterRole(scope, shape) === 'lane';
export const isSeparator = (scope: Scope, shape: Element) =>
	userValue(shape, VISIO_SEPARATOR_ROW) === '1' ||
	hasCategory(shape, 'Phase') ||
	masterRole(scope, shape) === 'separator';
const isList = (scope: Scope, shape: Element) =>
	userValue(shape, 'msvStructureType')?.toLowerCase() === 'list' ||
	['swimlane-list', 'phase-list'].includes(masterRole(scope, shape) ?? '');

/** The flowchart's Swimlane List and Phase List, when it has them (older files do not). */
export function listsOf(scope: Scope, frame: Element): { swim?: Element; phase?: Element } {
	const members = new Set(containerMembersOf(frame));
	const lists = topLevelShapes(scope.root).filter(
		(shape) => members.has(shapeIdOf(shape)) && isList(scope, shape),
	);
	const kind = (name: string) =>
		lists.find(
			(shape) =>
				userValue(shape, 'msvSDListRequiredCategories') === name ||
				masterRole(scope, shape) === (name === 'Swimlane' ? 'swimlane-list' : 'phase-list'),
		);
	const swim = kind('Swimlane');
	const phase = kind('Phase');
	return { ...(swim ? { swim } : {}), ...(phase ? { phase } : {}) };
}
/** The frame's members and, through its lists, the lanes and separators. */
export function frameMemberIds(scope: Scope, frame: Element): Set<string> {
	const { swim, phase } = listsOf(scope, frame);
	return new Set([
		...containerMembersOf(frame),
		...(swim ? listMembersOf(swim) : []),
		...(phase ? listMembersOf(phase) : []),
	]);
}
export const separatorsOf = (scope: Scope, frame: Element): Element[] => {
	const members = frameMemberIds(scope, frame);
	return topLevelShapes(scope.root).filter(
		(shape) => isSeparator(scope, shape) && members.has(shapeIdOf(shape)),
	);
};

/** The next `count` shape IDs after the page's highest. */
export function freeShapeIds(root: Element, count: number): string[] {
	let max = 0;
	for (const shape of Array.from(root.getElementsByTagName('Shape')))
		max = Math.max(max, Number(attribute(shape, 'ID')) || 0);
	return Array.from({ length: count }, (_, index) => String(max + 1 + index));
}

/**
 * Give the flowchart Visio's list structure: a Swimlane List and a Phase List that the frame
 * contains, the lanes and separators as their members. Flowcharts saved as plain containers (the
 * frame holding lanes and separators directly) are brought over the first time they are edited.
 */
export function ensureLists(
	scope: Scope,
	frame: Element,
	ids: { swim?: string; phase?: string } = {},
): { swim: Element; phase: Element } {
	const found = listsOf(scope, frame);
	if (found.swim && found.phase) {
		prepareFlowchart(scope, frame);
		return { swim: found.swim, phase: found.phase };
	}
	const options = optionsOf(frame);
	const lanes = lanesOf(scope, frame, options.direction);
	const bars = separatorsOf(scope, frame);
	const area = boundsOf(lanes) ?? boundsOf([frame])!;
	const fresh = freeShapeIds(scope.root, 2);
	const box = {
		x: (area.minX + area.maxX) / 2,
		y: (area.minY + area.maxY) / 2,
		width: area.maxX - area.minX,
		height: area.maxY - area.minY,
	};
	const horizontal = options.direction === 'horizontal';
	const make = (kind: 'swimlane' | 'phase', id: string) => {
		const list = createListShape(
			scope.root,
			scope.document,
			id,
			box,
			kind,
			(kind === 'swimlane') === horizontal ? 2 : 0,
			options.style,
		);
		frame.parentNode!.insertBefore(list, frame);
		return list;
	};
	const swim = found.swim ?? make('swimlane', ids.swim ?? fresh[0]!);
	const phase = found.phase ?? make('phase', ids.phase ?? fresh[1]!);
	const moved = new Set([...lanes, ...bars].map(shapeIdOf));
	writeContainerMembers(scope.root, frame, [
		...containerMembersOf(frame).filter((id) => !moved.has(id)),
		shapeIdOf(phase),
		shapeIdOf(swim),
	]);
	writeListMembers(scope.root, swim, lanes.map(shapeIdOf));
	writeListMembers(scope.root, phase, orderSeparators(bars, options.direction).map(shapeIdOf));
	setRelation(swim, VISIO_RELATIONSHIP_LIST_SIBLING, [shapeIdOf(phase)]);
	setRelation(phase, VISIO_RELATIONSHIP_LIST_SIBLING, [shapeIdOf(swim)]);
	return { swim, phase };
}
/** Separators in phase order: left to right in a horizontal flowchart, top to bottom otherwise. */
export function orderSeparators(
	bars: readonly Element[],
	direction: VisioSwimlaneDirection,
): Element[] {
	const key = (bar: Element) => {
		const box = boundsOf([bar]);
		return box ? (direction === 'horizontal' ? box.minX : -box.maxY) : 0;
	};
	return [...bars].sort((a, b) => key(a) - key(b));
}

/** The frame's lanes in stacking order: top to bottom, or left to right. */
export function lanesOf(
	scope: Scope,
	frame: Element,
	direction: VisioSwimlaneDirection,
): Element[] {
	const members = frameMemberIds(scope, frame);
	const lanes = topLevelShapes(scope.root).filter(
		(shape) => members.has(shapeIdOf(shape)) && isLane(scope, shape),
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
export function shift(scope: Scope, ids: Iterable<string>, dx: number, dy: number): void {
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

const NATIVE_CELLS = ['PinX', 'PinY', 'Width', 'Height', 'LocPinX', 'LocPinY', 'Angle'];
/** Give the instances of Visio's flowchart masters the cells the editor reads (size, pin). */
export function prepareFlowchart(scope: Scope, frame: Element): void {
	if (scope.masters.empty) return;
	const { swim, phase } = listsOf(scope, frame);
	const options = optionsOf(frame);
	const shapes = [
		frame,
		...(swim ? [swim] : []),
		...(phase ? [phase] : []),
		...lanesOf(scope, frame, options.direction),
		...separatorsOf(scope, frame),
	];
	for (const shape of shapes)
		if (masterRole(scope, shape)) materializeCells(scope.root, scope.masters, shape, NATIVE_CELLS);
}
const within = (shape: Element, instance: Element) => {
	for (let node: Node | null = instance; node; node = node.parentNode)
		if (node === shape) return true;
	return false;
};
/**
 * Give an instance of one of Visio's flowchart masters its box: the size first, then the master's
 * formulas recompute the caches of its parts (and its pin offset), then the pin follows.
 */
function placeInstance(scope: Scope, shape: Element, target: Box): void {
	const local = cells(shape);
	const set = (name: string, value: number) => local.get(name)!.setAttribute('V', String(value));
	if (!['Width', 'Height', 'PinX', 'PinY'].every((name) => local.has(name)))
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'A flowchart shape could not be resized.');
	set('Width', target.maxX - target.minX);
	set('Height', target.maxY - target.minY);
	const only = (instance: Element) => within(shape, instance);
	recalculateMasterInstances(scope.root, scope.masters, only);
	set('PinX', target.minX + numeric(local.get('LocPinX'), 0));
	set('PinY', target.minY + numeric(local.get('LocPinY'), 0));
	recalculateMasterInstances(scope.root, scope.masters, only);
}

/** Give a lane, separator or frame its box; a sheet the geometry editor refuses stops the edit. */
export function place(scope: Scope, shape: Element, target: Box): void {
	if (masterRole(scope, shape)) return placeInstance(scope, shape, target);
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
	const boxes = lanes.map((lane) => ({ id: shapeIdOf(lane), box: boundsOf([lane])! }));
	// A chart Visio built keeps the title area it has: the space between the frame and its lanes.
	const native = masterRole(scope, frame) === 'frame';
	const laneTop = Math.max(...boxes.map(({ box }) => box.maxY));
	const titleSize =
		native && options.title && Number.isFinite(laneTop)
			? Math.max(0, frameBox.maxY - laneTop)
			: undefined;
	const stacking = stackSwimlanes(options.direction, frameBox, options.title, boxes, titleSize);
	const carried = new Map(
		stacking.lanes.map((entry, index) => [index, [...containerClosure(scope.root, lanes[index]!)]]),
	);
	if (scope.contents === 'move') for (const ids of carried.values()) assertMovable(scope.root, ids);
	stacking.lanes.forEach((entry, index) => {
		scope.check();
		place(scope, lanes[index]!, entry.target);
		if (scope.contents === 'move') shift(scope, carried.get(index)!, entry.dx, entry.dy);
	});
	const bars = separatorsOf(scope, frame);
	for (const shape of bars) {
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
	setUserRow(frame, 'numLanes', String(lanes.length), '');
	// Visio's lists follow: the Swimlane List spans the lanes, both name their members in order.
	const { swim, phase } = listsOf(scope, frame);
	const area: Box = {
		minX: stacking.frame.minX,
		maxX: stacking.frame.maxX,
		maxY: stacking.top,
		minY: stacking.end,
	};
	for (const list of [swim, phase]) if (list) place(scope, list, area);
	if (swim) writeListMembers(scope.root, swim, lanes.map(shapeIdOf));
	if (phase)
		writeListMembers(scope.root, phase, orderSeparators(bars, options.direction).map(shapeIdOf));
}

/** Stencil instances move through the asynchronous instance path, so the caller must move them. */
export function assertMovable(root: Element, ids: readonly string[]): void {
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
