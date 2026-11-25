import { attribute, children } from './sheet';
import { VisioPackageError } from './package-common';
import { cells, numeric } from './edit-geometry-cells';
import { routeCrossings, VISIO_ROUTE_CLEARANCE } from './connector-route-avoid';
import type { VisioRouteBox, VisioRoutePoint } from './connector-route';
import { connectRows, proveConnector, topShape } from './edit-connector';
import { isVisioWalkGlue, visioGlueTrigger } from './edit-connector-glue';
import { connectorObstacles, obstacleBox } from './edit-connector-obstacles';
import { isRoutedConnector, rerouteConnector, routeOf } from './edit-connector-reroute';
import { effectiveCells, effectiveRows, isStencilConnector } from './edit-stencil-connector';
import { pageShapes } from './edit-instance-shape';
import { freeConnectorObstaclePath } from './connector-route-free';
import { layoutStencilConnector } from './edit-stencil-connector-layout';
import { executableCellFormula } from './cell-formula';

/**
 * The page-space vertices of a right-angle connector as it is saved, read from its cached
 * Geometry rows. Dynamic masters use the local pin, including the quarter-inch transform box
 * for level or plumb connectors. Local routed lines start at their begin point.
 */
export function connectorPolyline(shape: Element): VisioRoutePoint[] | undefined {
	const stencil = isStencilConnector(shape);
	const own = stencil ? effectiveCells(shape) : cells(shape);
	const read = (rows: Iterable<Map<string, Element>>): VisioRoutePoint[] => {
		const origin = stencil
			? {
					x: numeric(own.get('PinX')) - numeric(own.get('LocPinX')),
					y: numeric(own.get('PinY')) - numeric(own.get('LocPinY')),
				}
			: { x: numeric(own.get('BeginX')), y: numeric(own.get('BeginY')) };
		return [...rows].map((row) => ({
			x: origin.x + numeric(row.get('X')),
			y: origin.y + numeric(row.get('Y')),
		}));
	};
	if (stencil) {
		// ArcTo rows on right-angle routes are saved jumps; their chords are the route.
		const section = children(shape, 'Section').find((node) => attribute(node, 'N') === 'Geometry');
		if (
			children(section, 'Row').some((row) => {
				const type = attribute(row, 'T');
				return type && !['MoveTo', 'LineTo', 'ArcTo'].includes(type);
			})
		)
			return undefined;
		const rows = effectiveRows(shape, 'Geometry');
		return rows.size >= 2
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
	previous?: VisioRouteBox,
): readonly string[] {
	const root = roots.get(pageId);
	const moved = root && topShape(root, shapeId);
	const box = moved && obstacleBox(moved);
	if (!root || (!box && !previous)) return [];
	const adjacent = previous && {
		minX: previous.minX - VISIO_ROUTE_CLEARANCE - 1e-5,
		maxX: previous.maxX + VISIO_ROUTE_CLEARANCE + 1e-5,
		minY: previous.minY - VISIO_ROUTE_CLEARANCE - 1e-5,
		maxY: previous.maxY + VISIO_ROUTE_CLEARANCE + 1e-5,
	};
	const glued = new Set(connectRows(root).map((row) => attribute(row, 'FromSheet')));
	const pages = new Set<string>();
	let obstacles: VisioRouteBox[] | undefined;
	for (const shape of pageShapes(root)) {
		check();
		const id = attribute(shape, 'ID');
		if (!id || id === shapeId) continue;
		try {
			if (!isRoutedConnector(shape) || routeOf(shape) !== 'right-angle') continue;
			const own = isStencilConnector(shape) ? effectiveCells(shape) : cells(shape);
			const policy = numeric(own.get('ConFixedCode'), 0);
			if (policy === 1 || policy === 2) continue;
			const path = connectorPolyline(shape);
			if (!path) continue;
			const blocked = box && routeCrossings(path, [box]);
			if (!blocked && (!adjacent || glued.has(id) || !routeCrossings(path, [adjacent]))) continue;
			const glue = glued.has(id) ? proveConnector(root, id) : { shape, ends: {} };
			if (glue.ends.begin?.target === shapeId || glue.ends.end?.target === shapeId) continue;
			const free = !glue.ends.begin && !glue.ends.end;
			if (
				free &&
				['BeginX', 'BeginY', 'EndX', 'EndY'].some((name) => {
					const formula = executableCellFormula(attribute(own.get(name), 'F'));
					if (!formula) return false;
					const triggerName = name.startsWith('Begin') ? 'BegTrigger' : 'EndTrigger';
					const trigger = own.get(triggerName);
					const expression = attribute(trigger, 'F');
					return (
						!isVisioWalkGlue(formula) ||
						numeric(trigger) !== 1 ||
						(expression !== visioGlueTrigger(id) && expression !== '_XFTRIGGER(EventXFMod)')
					);
				})
			)
				continue;
			if (free) obstacles ??= connectorObstacles(root, new Set());
			const begin = { x: numeric(own.get('BeginX')), y: numeric(own.get('BeginY')) };
			const end = { x: numeric(own.get('EndX')), y: numeric(own.get('EndY')) };
			const cleared = !blocked && !!previous && !!routeCrossings([begin, end], [previous]);
			if (!blocked && !cleared) continue;
			const native =
				obstacles && free && isStencilConnector(shape)
					? freeConnectorObstaclePath(begin, end, obstacles, cleared)
					: undefined;
			const changed = native
				? layoutStencilConnector(root, pageId, shape, 'right-angle', native, check)
				: rerouteConnector(roots, pageId, glue, check, free && obstacles ? { obstacles } : {});
			for (const page of changed) pages.add(page);
		} catch (error) {
			if (!benign(error)) throw error;
		}
	}
	return [...pages];
}
