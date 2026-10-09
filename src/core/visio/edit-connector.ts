import { attribute, children } from './sheet';
import { fail } from './package-common';
import { transform } from './geometry';
import { executableCellFormula } from './cell-formula';
import { cells, numeric, setCell } from './edit-geometry-cells';
import { admitted, isLineSheet } from './edit-geometry-admission';
import { proveLocalLine, sameLineCoordinate } from './edit-line-move';
import { recalculateVisioCells, type VisioCellKey } from './edit-recalculate';
import type { VisioGeometryEdit } from './edit-commands';
import {
	VISIO_WALK_GLUE,
	isNativeGlueCell,
	visioConnectorGluePoints,
	visioGlueTrigger,
	visioGlueTriggerTarget,
	type VisioGlueBox,
	type VisioGluePoint,
} from './edit-connector-glue';

type LineCreate = Extract<VisioGeometryEdit, { type: 'create-line' }>;
type End = 'begin' | 'end';
const ENDS: readonly End[] = ['begin', 'end'];
const prefix = (end: End) => (end === 'begin' ? 'Begin' : 'End');
const triggerName = (end: End) => (end === 'begin' ? 'BegTrigger' : 'EndTrigger');
const UNSUPPORTED = 'Glued connections outside owned straight dynamic glue are unsupported.';

const connectRows = (root: Element) =>
	children(root, 'Connects').flatMap((container) => children(container, 'Connect'));
export function topShape(root: Element, shapeId: string): Element | undefined {
	return children(children(root, 'Shapes')[0], 'Shape').find(
		(shape) => attribute(shape, 'ID') === shapeId,
	);
}

/** A glue target's alignment box and transform, read from its own proven caches. */
function glueBox(root: Element, shapeId: string): VisioGlueBox {
	const shape = admitted(root, shapeId);
	const local = cells(shape);
	if (isLineSheet(local)) fail('UNSUPPORTED_GEOMETRY_EDIT', 'Connectors glue to 2D shapes only.');
	const width = numeric(local.get('Width')),
		height = numeric(local.get('Height'));
	const flip = (name: string) => numeric(local.get(name), 0) !== 0;
	return {
		width,
		height,
		transform: transform(
			numeric(local.get('PinX'), width / 2),
			numeric(local.get('PinY'), height / 2),
			numeric(local.get('LocPinX'), width / 2),
			numeric(local.get('LocPinY'), height / 2),
			numeric(local.get('Angle'), 0),
			flip('FlipX'),
			flip('FlipY'),
		),
	};
}

/** Endpoints of a new connector after its glued ends walk to their shapes' nearest sides. */
export function planConnector(root: Element, edit: LineCreate): LineCreate {
	if (!edit.connect) return edit;
	const site = (end: End): VisioGlueBox | VisioGluePoint => {
		const target = edit.connect?.[end];
		if (target !== undefined) return glueBox(root, target);
		return end === 'begin' ? { x: edit.beginX, y: edit.beginY } : { x: edit.endX, y: edit.endY };
	};
	const points = visioConnectorGluePoints(site('begin'), site('end'));
	if (Math.hypot(points.end.x - points.begin.x, points.end.y - points.begin.y) <= 0)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Glued connector endpoints coincide.');
	return {
		...edit,
		beginX: points.begin.x,
		beginY: points.begin.y,
		endX: points.end.x,
		endY: points.end.y,
	};
}

/** Turn a freshly drawn straight line into a native dynamic connector with Connect rows. */
export function glueNewConnector(root: Element, shape: Element, edit: LineCreate): void {
	if (!edit.connect) return;
	setCell(shape, 'ObjType', 2);
	setCell(shape, 'EndArrow', 4);
	const doc = root.ownerDocument!;
	let container = children(root, 'Connects')[0];
	for (const end of ENDS) {
		const target = edit.connect[end];
		setCell(shape, triggerName(end), target === undefined ? 0 : 2);
		if (target === undefined) continue;
		const local = cells(shape);
		local.get(triggerName(end))!.setAttribute('F', visioGlueTrigger(target));
		for (const axis of ['X', 'Y'])
			local.get(`${prefix(end)}${axis}`)!.setAttribute('F', VISIO_WALK_GLUE);
		if (!container) {
			container = doc.createElementNS(root.namespaceURI, 'Connects');
			const shapes = children(root, 'Shapes')[0]!;
			root.insertBefore(container, shapes.nextSibling);
		}
		const row = doc.createElementNS(root.namespaceURI, 'Connect');
		for (const [name, value] of [
			['FromSheet', edit.shapeId],
			['FromCell', `${prefix(end)}X`],
			['FromPart', end === 'begin' ? '9' : '12'],
			['ToSheet', target],
			['ToCell', 'PinX'],
			['ToPart', '3'],
		] as const)
			row.setAttribute(name, value);
		container.appendChild(row);
	}
}

interface ConnectorGlue {
	shape: Element;
	ends: Partial<Record<End, string>>;
}
/** Prove a connector's Connect rows and glue formulas are the owned dynamic-glue subset. */
function proveConnector(root: Element, connectorId: string): ConnectorGlue {
	const shape = topShape(root, connectorId);
	if (
		!shape ||
		shape.hasAttribute('Master') ||
		shape.hasAttribute('MasterShape') ||
		(attribute(shape, 'Type') ?? 'Shape') !== 'Shape' ||
		children(shape, 'Shapes').length
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', UNSUPPORTED);
	const local = cells(shape);
	const ends: Partial<Record<End, string>> = {};
	for (const row of connectRows(root)) {
		if (attribute(row, 'FromSheet') !== connectorId) continue;
		const end = ENDS.find((value) => attribute(row, 'FromCell') === `${prefix(value)}X`);
		const target = attribute(row, 'ToSheet');
		if (
			!end ||
			!target ||
			ends[end] !== undefined ||
			attribute(row, 'ToCell') !== 'PinX' ||
			attribute(row, 'ToPart') !== '3' ||
			visioGlueTriggerTarget(executableCellFormula(attribute(local.get(triggerName(end)), 'F'))) !==
				target ||
			['X', 'Y'].some(
				(axis) =>
					!isNativeGlueCell(
						`${prefix(end)}${axis}`,
						executableCellFormula(attribute(local.get(`${prefix(end)}${axis}`), 'F')) ?? '',
					),
			)
		)
			fail('UNSUPPORTED_GEOMETRY_EDIT', UNSUPPORTED);
		ends[end] = target;
	}
	return { shape, ends };
}

/** Classify a geometry target: connectors glued to it, and whether it is itself glued. */
export function glueParticipants(
	root: Element,
	shapeId: string,
): { connectors: ConnectorGlue[]; connector?: ConnectorGlue } {
	const connectorIds = new Set<string>();
	let own = false;
	for (const row of connectRows(root)) {
		if (attribute(row, 'ToSheet') === shapeId) connectorIds.add(attribute(row, 'FromSheet') ?? '');
		if (attribute(row, 'FromSheet') === shapeId) own = true;
	}
	if (own && connectorIds.size) fail('UNSUPPORTED_GEOMETRY_EDIT', UNSUPPORTED);
	return {
		connectors: [...connectorIds].map((id) => proveConnector(root, id)),
		...(own ? { connector: proveConnector(root, shapeId) } : {}),
	};
}

/** Detach connector ends: values stay, formulas, triggers and Connect rows go (Visio's unglue). */
export function unglueConnector(root: Element, glue: ConnectorGlue, ends: readonly End[]): void {
	const id = attribute(glue.shape, 'ID');
	for (const end of ends) {
		if (glue.ends[end] === undefined) continue;
		const local = cells(glue.shape);
		for (const axis of ['X', 'Y'])
			setCell(glue.shape, `${prefix(end)}${axis}`, numeric(local.get(`${prefix(end)}${axis}`)));
		setCell(glue.shape, triggerName(end), 0);
		for (const row of connectRows(root))
			if (attribute(row, 'FromSheet') === id && attribute(row, 'FromCell') === `${prefix(end)}X`)
				row.parentNode!.removeChild(row);
		delete glue.ends[end];
	}
	for (const container of children(root, 'Connects'))
		if (!children(container, 'Connect').length) root.removeChild(container);
}

/** Move glued ends back onto their shapes and recalculate the connector's derived caches. */
export function rerouteConnector(
	roots: ReadonlyMap<string, Element>,
	pageId: string,
	glue: ConnectorGlue,
	check: () => void,
): readonly string[] {
	const root = roots.get(pageId)!;
	const shape = glue.shape,
		shapeId = attribute(shape, 'ID')!;
	proveLocalLine(shape);
	const local = cells(shape);
	const site = (end: End): VisioGlueBox | VisioGluePoint => {
		const target = glue.ends[end];
		if (target !== undefined) return glueBox(root, target);
		return {
			x: numeric(local.get(`${prefix(end)}X`)),
			y: numeric(local.get(`${prefix(end)}Y`)),
		};
	};
	const points = visioConnectorGluePoints(site('begin'), site('end'));
	if (Math.hypot(points.end.x - points.begin.x, points.end.y - points.begin.y) <= 0)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Glued connector endpoints would coincide.');
	const changed: VisioCellKey[] = [];
	for (const end of ENDS)
		for (const axis of ['X', 'Y'] as const) {
			const name = `${prefix(end)}${axis}`,
				value = points[end][axis === 'X' ? 'x' : 'y'];
			if (sameLineCoordinate(numeric(local.get(name)), value)) continue;
			if (numeric(local.get(`Lock${prefix(end)}`), 0) !== 0)
				fail('EDIT_PROTECTED_CELL', `Lock${prefix(end)} prevents rerouting a glued connector.`);
			local.get(name)!.setAttribute('V', String(value));
			changed.push({ pageId, shapeId, cell: name });
		}
	if (!changed.length) return [];
	const pages = recalculateVisioCells(roots, changed, {
		check,
		lineEditShapes: new Set([shape]),
		glueShapes: new Set([shape]),
	});
	proveLocalLine(shape);
	return [...new Set([pageId, ...pages])];
}

/** Before deletion: drop removed connectors' rows and unglue retained connectors from removed shapes. */
export function releaseDeletedGlue(
	roots: ReadonlyMap<string, Element>,
	removed: ReadonlyMap<string, ReadonlySet<string>>,
): void {
	for (const [pageId, ids] of removed) {
		const root = roots.get(pageId);
		if (!root) continue;
		const touched = new Set<string>();
		for (const row of connectRows(root)) {
			const from = attribute(row, 'FromSheet') ?? '',
				to = attribute(row, 'ToSheet') ?? '';
			if (ids.has(from) || ids.has(to)) touched.add(from);
		}
		for (const connectorId of touched) {
			let glue: ConnectorGlue;
			try {
				glue = proveConnector(root, connectorId);
			} catch {
				// Foreign glue is not healed by guessing; the shared delete guard reports it.
				fail('EDIT_REFERENCED_DELETE', 'Shape participates in a Connect record.');
			}
			if (ids.has(connectorId)) {
				unglueConnector(root, glue, ENDS);
				continue;
			}
			unglueConnector(
				root,
				glue,
				ENDS.filter((end) => ids.has(glue.ends[end] ?? '')),
			);
		}
	}
}
