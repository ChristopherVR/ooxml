import { attribute, children } from './sheet';
import { VisioPackageError } from './package-common';
import { cells, numeric } from './edit-geometry-cells';
import { routeCrossings } from './connector-route-avoid';
import type { VisioRoutePoint } from './connector-route';
import { connectRows, proveConnector, topShape, type ConnectorGlue } from './edit-connector';
import { obstacleBox } from './edit-connector-obstacles';
import { rerouteConnector, routeOf } from './edit-connector-reroute';
import { effectiveCells, effectiveRows, isStencilConnector } from './edit-stencil-connector';

/**
 * The page-space vertices of a right-angle connector as it is saved, read from its cached
 * Geometry rows (relative to its begin point, the form both this editor and Visio's Dynamic
 * connector use), or nothing when the rows are anything but one MoveTo and LineTo runs.
 */
export function connectorPolyline(shape: Element): VisioRoutePoint[] | undefined {
	const stencil = isStencilConnector(shape);
	const own = stencil ? effectiveCells(shape) : cells(shape);
	const read = (rows: Iterable<Map<string, Element>>): VisioRoutePoint[] => {
		const origin = { x: numeric(own.get('BeginX')), y: numeric(own.get('BeginY')) };
		return [...rows].map((row) => ({
			x: origin.x + numeric(row.get('X')),
			y: origin.y + numeric(row.get('Y')),
		}));
	};
	if (stencil) {
		// The master's rows are proven to be a MoveTo followed by LineTo rows.
		const rows = effectiveRows(shape, 'Geometry');
		return [...rows.keys()].length >= 2
			? read([...rows.entries()].sort((a, b) => a[0] - b[0]).map(([, row]) => row))
			: undefined;
	}
	const sections = children(shape, 'Section').filter((node) => attribute(node, 'N') === 'Geometry');
	if (sections.length !== 1) return undefined;
	const rows = children(sections[0], 'Row');
	if (
		rows.length < 2 ||
		attribute(rows[0], 'T') !== 'MoveTo' ||
		rows.slice(1).some((row) => attribute(row, 'T') !== 'LineTo')
	)
		return undefined;
	return read(rows.map((row) => cells(row)));
}

const benign = (error: unknown) =>
	error instanceof VisioPackageError && !error.code.startsWith('LIMIT_');

/**
 * Visio lays a right-angle connector out again when a placeable shape is moved, resized or turned
 * onto its path. Every right-angle connector of the page whose saved path now runs through the
 * shape's box, and that is not glued to the shape, is routed again around it. A connector this
 * editor cannot lay out again is left as it is: it is not part of the edit that was asked for.
 */
export function rerouteBlockedConnectors(
	roots: ReadonlyMap<string, Element>,
	pageId: string,
	shapeId: string,
	check: () => void,
): readonly string[] {
	const root = roots.get(pageId);
	const moved = root && topShape(root, shapeId);
	const box = moved && obstacleBox(moved);
	if (!root || !box) return [];
	const pages = new Set<string>();
	for (const id of new Set(connectRows(root).map((row) => attribute(row, 'FromSheet') ?? ''))) {
		check();
		if (!id || id === shapeId) continue;
		try {
			const shape = topShape(root, id);
			if (!shape || routeOf(shape) !== 'right-angle') continue;
			const path = connectorPolyline(shape);
			if (!path || !routeCrossings(path, [box])) continue;
			const glue: ConnectorGlue = proveConnector(root, id);
			if (glue.ends.begin?.target === shapeId || glue.ends.end?.target === shapeId) continue;
			for (const page of rerouteConnector(roots, pageId, glue, check)) pages.add(page);
		} catch (error) {
			if (!benign(error)) throw error;
		}
	}
	return [...pages];
}
