import { attribute, children } from './sheet';
import { fail } from './package-common';
import { executableCellFormula } from './cell-formula';
import { cells, numeric } from './edit-geometry-cells';
import { isLineSheet } from './edit-geometry-admission';
import { proveLocalLine, sameLineCoordinate } from './edit-line-move';
import type { VisioConnectorEdit, VisioConnectorRoute } from './edit-connector-commands';
import { addConnectionPoint, deleteConnectionPoint } from './edit-connection-points';
import type { VisioGeometryEdit } from './edit-commands';
import {
	ENDS,
	chooseSites,
	connectorRoute,
	layoutConnectorShape,
	prefix,
	setRouteCells,
	type ConnectorSite,
	type End,
} from './edit-connector-layout';
import {
	glueEnd,
	glueSites,
	topShape,
	unglueConnector,
	proveConnector,
	type ConnectorGlue,
} from './edit-connector';
import type { VisioRoutePoint } from './connector-route';

const canonical = (shape: Element, name: string, expected: string) =>
	executableCellFormula(attribute(cells(shape).get(name), 'F'))
		?.replace(/\s+/g, '')
		.toLowerCase() === expected.toLowerCase();

/** A right-angle or curved connector this editor wrote (Visio's dynamic-connector cell form). */
export function isRoutedConnector(shape: Element): boolean {
	return isLineSheet(cells(shape)) && connectorRoute(shape) !== 'straight';
}
/** Prove a connector's transform is the owned form of its route before it is laid out again. */
export function proveConnectorShape(shape: Element, route = connectorRoute(shape)): void {
	if (route === 'straight') {
		proveLocalLine(shape);
		return;
	}
	const local = cells(shape);
	const value = (name: string) => numeric(local.get(name));
	if (
		!canonical(shape, 'PinX', '(BeginX+EndX)/2') ||
		!canonical(shape, 'PinY', '(BeginY+EndY)/2') ||
		!canonical(shape, 'Width', 'EndX-BeginX') ||
		!canonical(shape, 'Height', 'EndY-BeginY') ||
		numeric(local.get('Angle'), 0) !== 0 ||
		numeric(local.get('FlipX'), 0) !== 0 ||
		numeric(local.get('FlipY'), 0) !== 0 ||
		!sameLineCoordinate(value('Width'), value('EndX') - value('BeginX')) ||
		!sameLineCoordinate(value('Height'), value('EndY') - value('BeginY')) ||
		children(shape, 'Section').filter((node) => attribute(node, 'N') === 'Geometry').length !== 1
	)
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Only right-angle and curved connectors drawn here can be rerouted.',
		);
}

/**
 * Lay a connector out again: glued ends move onto their shapes (nearest side midpoints) or
 * connection points, unglued ends stay (or take `free`), and the route is redrawn.
 */
export function rerouteConnector(
	roots: ReadonlyMap<string, Element>,
	pageId: string,
	glue: ConnectorGlue,
	check: () => void,
	options: { free?: Partial<Record<End, VisioRoutePoint>>; route?: VisioConnectorRoute } = {},
): readonly string[] {
	const root = roots.get(pageId)!;
	const shape = glue.shape;
	const current = connectorRoute(shape);
	proveConnectorShape(shape, current);
	const local = cells(shape);
	const sites = (end: End): ConnectorSite[] => {
		const target = glue.ends[end];
		if (target) return glueSites(root, target);
		return [
			{
				point: options.free?.[end] ?? {
					x: numeric(local.get(`${prefix(end)}X`)),
					y: numeric(local.get(`${prefix(end)}Y`)),
				},
			},
		];
	};
	const chosen = chooseSites(sites('begin'), sites('end'));
	const route = options.route ?? current;
	if (
		route === 'straight' &&
		current === 'straight' &&
		ENDS.every(
			(end) =>
				sameLineCoordinate(numeric(local.get(`${prefix(end)}X`)), chosen[end].point.x) &&
				sameLineCoordinate(numeric(local.get(`${prefix(end)}Y`)), chosen[end].point.y),
		)
	)
		return [];
	if (options.route) setRouteCells(shape, options.route);
	const pages = layoutConnectorShape(roots, pageId, shape, route, chosen, check);
	proveConnectorShape(shape, route);
	return pages;
}

const connectorShape = (root: Element, shapeId: string): Element => {
	const shape = topShape(root, shapeId);
	if (!shape || !isLineSheet(cells(shape)) || numeric(cells(shape).get('ObjType'), 0) !== 2)
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Only connectors drawn with the Connector tool can be glued or rerouted.',
		);
	return shape;
};

/** Move, endpoint and transform edits of a right-angle or curved connector. */
export function editRoutedConnector(
	roots: ReadonlyMap<string, Element>,
	edit: VisioGeometryEdit,
	check: () => void,
): readonly string[] {
	const root = roots.get(edit.pageId)!;
	const glue = proveConnector(root, edit.shapeId);
	const local = cells(glue.shape);
	const point = (end: End) => ({
		x: numeric(local.get(`${prefix(end)}X`)),
		y: numeric(local.get(`${prefix(end)}Y`)),
	});
	if (edit.type === 'move-shape') {
		const dx = edit.x - numeric(local.get('PinX')),
			dy = edit.y - numeric(local.get('PinY'));
		if (dx === 0 && dy === 0) return [];
		for (const [delta, lock] of [
			[dx, 'LockMoveX'],
			[dy, 'LockMoveY'],
		] as const)
			if (delta !== 0 && numeric(local.get(lock), 0) !== 0)
				fail('EDIT_PROTECTED_CELL', `${lock} prevents this operation.`);
		const free = Object.fromEntries(
			ENDS.map((end) => [end, { x: point(end).x + dx, y: point(end).y + dy }]),
		);
		unglueConnector(root, glue, ENDS);
		return rerouteConnector(roots, edit.pageId, glue, check, { free });
	}
	if (edit.type === 'move-line-endpoint') {
		const end = edit.endpoint === 'begin' ? 'begin' : 'end';
		unglueConnector(root, glue, [end]);
		return rerouteConnector(roots, edit.pageId, glue, check, {
			free: { [end]: { x: edit.x, y: edit.y } },
		});
	}
	return fail(
		'UNSUPPORTED_GEOMETRY_EDIT',
		'Right-angle and curved connectors are reshaped by their ends and route, not resized, rotated or flipped.',
	);
}

/** Glue one end of an existing connector (Visio's re-glue by dragging an endpoint). */
export function glueExistingConnector(
	roots: ReadonlyMap<string, Element>,
	edit: Extract<VisioGeometryEdit, { type: 'glue-connector' }>,
	check: () => void,
): readonly string[] {
	const root = roots.get(edit.pageId)!;
	const shape = connectorShape(root, edit.shapeId);
	const glue = proveConnector(root, edit.shapeId);
	const end = edit.endpoint === 'begin' ? 'begin' : 'end';
	const other = glue.ends[end === 'begin' ? 'end' : 'begin'];
	if (other?.target === edit.target)
		fail('INVALID_EDIT', 'A connector cannot glue both ends to the same shape.');
	const target = {
		target: edit.target,
		...(edit.point === undefined ? {} : { point: edit.point }),
	};
	glueSites(root, target);
	proveConnectorShape(shape);
	unglueConnector(root, glue, [end]);
	glueEnd(root, shape, end, target);
	glue.ends[end] = target;
	const pages = rerouteConnector(roots, edit.pageId, glue, check);
	return pages.length ? pages : [edit.pageId];
}

/** Design > Connectors: redraw a connector as right-angle, straight or curved. */
export function setConnectorRoute(
	roots: ReadonlyMap<string, Element>,
	edit: Extract<VisioGeometryEdit, { type: 'set-connector-route' }>,
	check: () => void,
): readonly string[] {
	const root = roots.get(edit.pageId)!;
	const shape = connectorShape(root, edit.shapeId);
	if (connectorRoute(shape) === edit.route) return [];
	const glue = proveConnector(root, edit.shapeId);
	const pages = rerouteConnector(roots, edit.pageId, glue, check, { route: edit.route });
	return pages.length ? pages : [edit.pageId];
}

/** Connection-point, glue and route commands. */
export function applyConnectorEdit(
	roots: ReadonlyMap<string, Element>,
	root: Element,
	edit: VisioConnectorEdit,
	check: () => void,
): readonly string[] {
	switch (edit.type) {
		case 'add-connection-point':
			return addConnectionPoint(root, edit);
		case 'delete-connection-point':
			return deleteConnectionPoint(root, edit);
		case 'glue-connector':
			return glueExistingConnector(roots, edit, check);
		case 'set-connector-route':
			return setConnectorRoute(roots, edit, check);
	}
}
