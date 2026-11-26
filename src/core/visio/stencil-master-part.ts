import { NS, parseXml } from '../xml/index';
import { setCell } from './edit-geometry-cells';
import { appendEllipseGeometry } from './edit-ellipse-geometry';
import { applyShapeCreationStyles, outlineGeometrySections } from './edit-shape-create';
import { visioOutlineShape } from './stencil-shapes';
import { appendRelativeConnectionPoint } from './edit-connection-points';
import { writePropertyRow } from './edit-shape-data';
import type { VisioBuiltInMaster } from './stencil-masters';

/** The one shape of a built-in master; instances name it through their `Master` attribute. */
export const STENCIL_MASTER_SHAPE_ID = '5';
const PROTECTION = [
	'LockWidth',
	'LockHeight',
	'LockMoveX',
	'LockMoveY',
	'LockAspect',
	'LockDelete',
	'LockRotate',
	'LockTextEdit',
	'LockFormat',
	'LockGroup',
];

function textCell(owner: Element, name: string, value: string): void {
	const cell = owner.ownerDocument!.createElementNS(owner.namespaceURI, 'Cell');
	cell.setAttribute('N', name);
	cell.setAttribute('V', value);
	const before = Array.from(owner.childNodes).find(
		(child) => child.nodeType === 1 && (child as Element).localName !== 'Cell',
	);
	owner.insertBefore(cell, before ?? null);
}

/**
 * The `<Master>` entry of a built-in master for masters.xml: its names, its identity and a master
 * page as large as the shape. `IconUpdate` asks Visio to draw the icon itself; none is stored.
 */
export function stencilMasterEntry(
	masters: Element,
	builtIn: Pick<VisioBuiltInMaster, 'uniqueId' | 'baseId' | 'layer'> & {
		master: Pick<VisioBuiltInMaster['master'], 'size'>;
	},
	id: string,
	name: string,
	relationshipId: string,
): Element {
	const doc = masters.ownerDocument!;
	const node = (parent: Element, tag: string) =>
		parent.appendChild(doc.createElementNS(masters.namespaceURI, tag)) as Element;
	const master = node(masters, 'Master');
	for (const [attribute, value] of [
		['ID', id],
		['NameU', name],
		['IsCustomNameU', '1'],
		['Name', name],
		['IsCustomName', '1'],
		['Prompt', 'Drag the shape onto the drawing page.'],
		['IconSize', '1'],
		['AlignName', '2'],
		['MatchByName', '0'],
		['IconUpdate', '1'],
		['UniqueID', builtIn.uniqueId],
		['BaseID', builtIn.baseId],
		['PatternFlags', '0'],
		['Hidden', '0'],
		['MasterType', '2'],
	] as const)
		master.setAttribute(attribute, value);
	const sheet = node(master, 'PageSheet');
	for (const style of ['LineStyle', 'FillStyle', 'TextStyle']) sheet.setAttribute(style, '0');
	const { width, height } = builtIn.master.size;
	for (const [cell, value] of [
		['PageWidth', width],
		['PageHeight', height],
		['PageScale', 1],
		['DrawingScale', 1],
		['DrawingSizeType', 4],
		['DrawingScaleType', 0],
	] as const)
		setCell(sheet, cell, value);
	if (builtIn.layer) {
		const section = node(sheet, 'Section');
		section.setAttribute('N', 'Layer');
		const row = node(section, 'Row');
		row.setAttribute('IX', '0');
		for (const [cell, value] of [
			['Name', builtIn.layer],
			['Color', '255'],
			['Status', '0'],
			['Visible', '1'],
			['Print', '1'],
			['Active', '0'],
			['Lock', '0'],
			['Snap', '1'],
			['Glue', '1'],
			['NameUniv', builtIn.layer],
			['ColorTrans', '0'],
		] as const) {
			const layerCell = node(row, 'Cell');
			layerCell.setAttribute('N', cell);
			layerCell.setAttribute('V', value);
		}
	}
	node(master, 'Rel').setAttributeNS(NS.r, 'r:id', relationshipId);
	return master;
}

/**
 * The master part of a built-in master: one shape in the drawing's default styles, sized as the
 * master drops, whose pin, text block and outline follow its width and height. The outline is
 * this package's own geometry (relative rows, or the native ellipse row), never a Visio master's.
 */
export function stencilMasterContents(
	namespace: string,
	document: Element,
	builtIn: VisioBuiltInMaster,
): Element {
	const root = parseXml(`<MasterContents xmlns="${namespace}"/>`).documentElement;
	root.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
	const doc = root.ownerDocument!;
	const shapes = root.appendChild(doc.createElementNS(namespace, 'Shapes')) as Element;
	const shape = shapes.appendChild(doc.createElementNS(namespace, 'Shape')) as Element;
	shape.setAttribute('ID', STENCIL_MASTER_SHAPE_ID);
	shape.setAttribute('Type', 'Shape');
	applyShapeCreationStyles(shape, document);
	const { width, height } = builtIn.master.size;
	for (const [name, value, formula] of [
		['PinX', width / 2],
		['PinY', height / 2],
		['Width', width],
		['Height', height],
		['LocPinX', width / 2, 'Width*0.5'],
		['LocPinY', height / 2, 'Height*0.5'],
		['Angle', 0],
		['FlipX', 0],
		['FlipY', 0],
		['ResizeMode', 0],
		['ObjType', 1],
	] as const)
		setCell(shape, name, value, formula);
	// Lock the aspect of the equal-sided shapes, and write the remaining protection explicitly.
	for (const name of PROTECTION)
		setCell(
			shape,
			name,
			Number(name === 'LockAspect' && ['square', 'circle'].includes(builtIn.master.id)),
		);
	if (builtIn.layer) textCell(shape, 'LayerMember', '0');
	const outline =
		builtIn.master.shape === 'ellipse' ? undefined : visioOutlineShape(builtIn.master.shape);
	const connection = doc.createElementNS(namespace, 'Section');
	connection.setAttribute('N', 'Connection');
	shape.appendChild(connection);
	for (const [index, point] of cardinalPoints(outline?.paths[0]).entries())
		appendRelativeConnectionPoint(connection, index, builtIn.master.size, point);
	writePropertyRow(shape, 'ShapeKind', {
		label: 'Shape type',
		prompt: '',
		type: 'string',
		format: '',
		value: builtIn.master.name,
	});
	if (!outline) appendEllipseGeometry(shape, width, height);
	else {
		// A fixed corner radius, as a drawn rounded rectangle keeps when it is resized.
		if (outline.rounding) setCell(shape, 'Rounding', outline.rounding * Math.min(width, height));
		for (const section of outlineGeometrySections(shape, outline)) shape.appendChild(section);
	}
	return root;
}

/** Farthest intersections with the outline at its horizontal and vertical midlines. */
function cardinalPoints(path: readonly (readonly [number, number])[] | undefined) {
	const extent = (axis: 0 | 1, high: boolean) => {
		const other = axis === 0 ? 1 : 0;
		let edge = high ? -Infinity : Infinity;
		let previous = path?.at(-1);
		for (const point of path ?? []) {
			if (previous && previous[other] > 0.5 !== point[other] > 0.5) {
				const fraction = (0.5 - previous[other]) / (point[other] - previous[other]);
				const at = previous[axis] + fraction * (point[axis] - previous[axis]);
				edge = high ? Math.max(edge, at) : Math.min(edge, at);
			}
			previous = point;
		}
		return Number.isFinite(edge) ? edge : Number(high);
	};
	return [
		{ x: extent(0, false), y: 0.5 },
		{ x: extent(0, true), y: 0.5 },
		{ x: 0.5, y: extent(1, false) },
		{ x: 0.5, y: extent(1, true) },
	];
}
