import { VisioPackageError } from './package-common';
import { cells, numeric } from './edit-geometry-cells';
import { visioShapeBounds } from './edit-group';
import { applyGeometryEdit } from './edit-geometry';
import {
	VISIO_CONTAINER_HEADING_ROW,
	VISIO_CONTAINER_MARGIN,
	VISIO_CONTAINER_RESIZE_ROW,
	VISIO_CONTAINER_MARGIN_ROW,
	type VisioReconcileContainersEdit,
} from './edit-diagram-parts-commands';
import {
	containerClosure,
	containerMembersOf,
	isContainerSheet,
	shapeIdOf,
	topLevelShapes,
	userValue,
	writeContainerMembers,
} from './edit-container-members';

/** Visio's msvSDContainerResize: how a container follows its members. */
export const CONTAINER_RESIZE_NONE = 0;
export const CONTAINER_RESIZE_EXPAND = 1;
export const CONTAINER_RESIZE_TIGHT = 2;

export interface Box {
	minX: number;
	minY: number;
	maxX: number;
	maxY: number;
}
const EPSILON = 1e-6;

export function boundsOf(shapes: readonly Element[]): Box | undefined {
	try {
		return shapes.length ? visioShapeBounds(shapes) : undefined;
	} catch {
		// A shape without usable local transform caches (a stencil instance) has no known bounds.
		return undefined;
	}
}
const area = (box: Box) => (box.maxX - box.minX) * (box.maxY - box.minY);
const number = (text: string | undefined, fallback: number) => {
	const value = text === undefined || text === '' ? NaN : Number(text);
	return Number.isFinite(value) ? value : fallback;
};
const unrotated = (shape: Element) =>
	numeric(cells(shape).get('Angle'), 0) === 0 &&
	numeric(cells(shape).get('FlipX'), 0) === 0 &&
	numeric(cells(shape).get('FlipY'), 0) === 0;

/**
 * Give a top-level unrotated sheet the page box `target`: resize with the lower-left corner
 * fixed, then slide it to the wanted corner. False when the geometry editor does not admit it.
 */
export function placeVisioBox(
	roots: ReadonlyMap<string, Element>,
	document: Element,
	pageId: string,
	container: Element,
	target: Box,
	check: () => void,
): boolean {
	try {
		// Resize with the lower-left corner fixed, then slide the sheet to the wanted corner.
		applyGeometryEdit(
			roots,
			document,
			{
				type: 'resize-shape',
				pageId,
				shapeId: shapeIdOf(container),
				width: target.maxX - target.minX,
				height: target.maxY - target.minY,
				anchor: { x: 0, y: 0 },
			},
			check,
		);
		const moved = boundsOf([container])!;
		applyGeometryEdit(
			roots,
			document,
			{
				type: 'move-shape',
				pageId,
				shapeId: shapeIdOf(container),
				x: numeric(cells(container).get('PinX')) + (target.minX - moved.minX),
				y: numeric(cells(container).get('PinY')) + (target.minY - moved.minY),
			},
			check,
		);
	} catch (error) {
		// A container the geometry editor does not admit keeps its size; membership still stands.
		if (!(error instanceof VisioPackageError)) throw error;
		return false;
	}
	return true;
}

/**
 * Grow (or, when always-fit, shrink) a container so it surrounds its members with its margin and
 * heading allowance. Returns whether the container changed. Only an unrotated ordinary container
 * is resized; any other is left alone.
 */
export function fitVisioContainer(
	roots: ReadonlyMap<string, Element>,
	document: Element,
	pageId: string,
	container: Element,
	check: () => void,
): boolean {
	const root = roots.get(pageId)!;
	const mode = number(userValue(container, VISIO_CONTAINER_RESIZE_ROW), CONTAINER_RESIZE_EXPAND);
	const memberIds = new Set(containerMembersOf(container));
	const members = topLevelShapes(root).filter((shape) => memberIds.has(shapeIdOf(shape)));
	const need = boundsOf(members);
	const have = boundsOf([container]);
	if (mode === CONTAINER_RESIZE_NONE || !need || !have || !unrotated(container)) return false;
	const margin = number(userValue(container, VISIO_CONTAINER_MARGIN_ROW), VISIO_CONTAINER_MARGIN);
	const heading = number(userValue(container, VISIO_CONTAINER_HEADING_ROW), 0);
	const wanted: Box = {
		minX: need.minX - margin,
		minY: need.minY - margin,
		maxX: need.maxX + margin,
		maxY: need.maxY + margin + heading,
	};
	const target: Box =
		mode === CONTAINER_RESIZE_TIGHT
			? wanted
			: {
					minX: Math.min(have.minX, wanted.minX),
					minY: Math.min(have.minY, wanted.minY),
					maxX: Math.max(have.maxX, wanted.maxX),
					maxY: Math.max(have.maxY, wanted.maxY),
				};
	if (
		(['minX', 'minY', 'maxX', 'maxY'] as const).every(
			(side) => Math.abs(target[side] - have[side]) < EPSILON,
		)
	)
		return false;
	return placeVisioBox(roots, document, pageId, container, target, check);
}

/**
 * After shapes moved or resized: a shape fully inside a container joins it (the innermost one), a
 * member whose centre left its container leaves it, and every container whose members changed
 * grows to fit (and so do the containers that hold it). Shapes that moved together with their
 * container keep their membership. Returns whether the page changed.
 */
export function reconcileVisioContainers(
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: VisioReconcileContainersEdit,
	check: () => void,
): boolean {
	check();
	const root = roots.get(edit.pageId)!;
	const shapes = topLevelShapes(root);
	const byId = new Map(shapes.map((shape) => [shapeIdOf(shape), shape]));
	const containers = shapes.filter(isContainerSheet);
	if (!containers.length) return false;
	const moved = new Set(edit.shapeIds.filter((id) => byId.has(id)));
	const touched = new Set<Element>();
	let changed = false;
	const boxes = new Map(containers.map((c) => [c, boundsOf([c])]));
	for (const id of moved) {
		const shape = byId.get(id)!;
		const box = boundsOf([shape]);
		if (!box) continue;
		const centre = { x: (box.minX + box.maxX) / 2, y: (box.minY + box.maxY) / 2 };
		const mine = isContainerSheet(shape) ? containerClosure(root, shape) : new Set<string>();
		const candidates = containers.filter((c) => c !== shape && !moved.has(shapeIdOf(c)));
		let inner: Element | undefined;
		for (const container of candidates) {
			const frame = boxes.get(container);
			if (!frame || mine.has(shapeIdOf(container))) continue;
			const members = containerMembersOf(container);
			const inside =
				box.minX >= frame.minX - EPSILON &&
				box.maxX <= frame.maxX + EPSILON &&
				box.minY >= frame.minY - EPSILON &&
				box.maxY <= frame.maxY + EPSILON;
			const centred =
				centre.x >= frame.minX &&
				centre.x <= frame.maxX &&
				centre.y >= frame.minY &&
				centre.y <= frame.maxY;
			if (members.includes(id) && !centred) {
				writeContainerMembers(
					root,
					container,
					members.filter((m) => m !== id),
				);
				touched.add(container);
				changed = true;
			} else if (
				inside &&
				!members.includes(id) &&
				(!inner || area(frame) < area(boxes.get(inner)!))
			)
				inner = container;
			else if (members.includes(id)) touched.add(container);
		}
		if (inner) {
			writeContainerMembers(root, inner, [...containerMembersOf(inner), id]);
			touched.add(inner);
			changed = true;
		}
	}
	// Grow outwards: a fitted container may itself be a member of another one.
	const queue = [...touched];
	for (let step = 0; queue.length && step < 1000; step++) {
		const container = queue.shift()!;
		if (!fitVisioContainer(roots, document, edit.pageId, container, check)) continue;
		changed = true;
		for (const parent of containers)
			if (
				containerMembersOf(parent).includes(shapeIdOf(container)) &&
				!moved.has(shapeIdOf(parent))
			)
				queue.push(parent);
	}
	return changed;
}
