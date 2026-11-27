import type { VisioRouteBox, VisioRoutePoint } from './connector-route';
import {
	spaceVisioRoute,
	visioConnectorSpacing,
	type VisioConnectorSpacing,
	type VisioSpacedRoute,
} from './connector-spacing';
import { connectRows, type ConnectorGlue } from './edit-connector';
import { cells, numeric } from './edit-geometry-cells';
import { isLineSheet } from './edit-geometry-admission';
import { effectiveCells, isStencilConnector } from './edit-stencil-connector';
import type { VisioPackage } from './package';
import { related, visioXml } from './parts';
import { attribute, children } from './sheet';

/** Each page root's spacing policy, read from its PageSheet before the edits run. */
const policies = new WeakMap<Element, VisioConnectorSpacing>();

const pageCell = (sheet: Element | undefined, name: string): number | undefined => {
	const cell = children(sheet, 'Cell').find((node) => attribute(node, 'N') === name);
	const value = Number(attribute(cell, 'V'));
	return cell && Number.isFinite(value) ? value : undefined;
};

/**
 * Read Page Setup > Layout and Routing's Separate and spacing cells (LineAdjustFrom, LineToLineX,
 * LineToLineY) of every page being edited, so connectors routed in this transaction are spaced
 * apart as Visio spaces them. Pages that store none of the cells take Visio's defaults.
 */
export async function registerConnectorSpacing(
	pkg: VisioPackage,
	roots: ReadonlyMap<string, Element>,
	check: () => void,
): Promise<void> {
	const documentPart = await related(pkg, '', 'document');
	const pagesPart = documentPart && (await related(pkg, documentPart, 'pages'));
	if (!pagesPart) return;
	const list = await visioXml(pkg, pagesPart, 'Pages');
	for (const page of children(list, 'Page')) {
		check();
		const root = roots.get(attribute(page, 'ID') ?? '');
		if (!root) continue;
		const sheet = children(page, 'PageSheet')[0];
		policies.set(
			root,
			visioConnectorSpacing(
				pageCell(sheet, 'LineAdjustFrom'),
				pageCell(sheet, 'LineToLineX'),
				pageCell(sheet, 'LineToLineY'),
			),
		);
	}
}

/** A connector's drawn route in page coordinates, or undefined when it is not a plain polyline. */
function drawnRoute(shape: Element): VisioRoutePoint[] | undefined {
	const local = isStencilConnector(shape) ? effectiveCells(shape) : cells(shape);
	if (!isLineSheet(local)) return undefined;
	if (numeric(local.get('Angle'), 0) !== 0 || numeric(local.get('FlipX'), 0) !== 0)
		return undefined;
	if (numeric(local.get('FlipY'), 0) !== 0) return undefined;
	const sections = children(shape, 'Section').filter((node) => attribute(node, 'N') === 'Geometry');
	if (sections.length !== 1) return undefined;
	const originX = numeric(local.get('PinX'), NaN) - numeric(local.get('LocPinX'), NaN),
		originY = numeric(local.get('PinY'), NaN) - numeric(local.get('LocPinY'), NaN);
	if (!Number.isFinite(originX) || !Number.isFinite(originY)) return undefined;
	const points: VisioRoutePoint[] = [];
	for (const row of children(sections[0], 'Row')) {
		const type = attribute(row, 'T');
		if (type !== 'MoveTo' && type !== 'LineTo') return undefined;
		const value = (name: string) =>
			Number(
				attribute(
					children(row, 'Cell').find((cell) => attribute(cell, 'N') === name),
					'V',
				),
			);
		const x = value('X'),
			y = value('Y');
		if (!Number.isFinite(x) || !Number.isFinite(y)) return undefined;
		points.push({ x: originX + x, y: originY + y });
	}
	return points.length >= 2 ? points : undefined;
}

/** The other connectors on the page with the shapes they are glued to. */
function pageRoutes(root: Element, except: Element): VisioSpacedRoute[] {
	const glued = new Map<string, Set<string>>();
	for (const row of connectRows(root)) {
		const from = attribute(row, 'FromSheet') ?? '';
		if (!/^(Begin|End)X$/.test(attribute(row, 'FromCell') ?? '')) continue;
		const set = glued.get(from) ?? new Set<string>();
		set.add(attribute(row, 'ToSheet') ?? '');
		glued.set(from, set);
	}
	const routes: VisioSpacedRoute[] = [];
	for (const shape of children(children(root, 'Shapes')[0], 'Shape')) {
		if (shape === except) continue;
		const local = isStencilConnector(shape) ? effectiveCells(shape) : cells(shape);
		if (numeric(local.get('ObjType'), 0) !== 2) continue;
		const points = drawnRoute(shape);
		if (points)
			routes.push({ points, targets: glued.get(attribute(shape, 'ID') ?? '') ?? new Set() });
	}
	return routes;
}

/**
 * Space a right-angle connector's route apart from the connectors it would lie on (see
 * {@link spaceVisioRoute}), with the page's registered policy; unregistered pages are unchanged.
 */
export function spaceConnectorRoute(
	root: Element,
	glue: ConnectorGlue,
	vertices: readonly VisioRoutePoint[],
	obstacles: readonly VisioRouteBox[],
): VisioRoutePoint[] {
	const policy = policies.get(root);
	if (!policy || policy.separate === 'none') return [...vertices];
	const targets = new Set(
		[glue.ends.begin?.target, glue.ends.end?.target].filter(
			(target): target is string => target !== undefined,
		),
	);
	return spaceVisioRoute(vertices, targets, pageRoutes(root, glue.shape), policy, obstacles);
}
