import { attribute, children } from './sheet';
import { cells, numeric, setCell } from './edit-geometry-cells';
import {
	applyShapeCreationStyles,
	createShape,
	outlineGeometrySections,
} from './edit-shape-create';
import { node, placeBox, textNode, userSection, valueCell } from './edit-diagram-parts';
import { VISIO_SEPARATOR_ROW, VISIO_SWIMLANE_LANE_ROW } from './shape-structure';
import {
	VISIO_CONTAINER_MARGIN_ROW,
	VISIO_CONTAINER_RESIZE_ROW,
	type VisioPartBox,
} from './edit-diagram-parts-commands';
import {
	VISIO_SEPARATOR_THICKNESS,
	VISIO_SWIMLANE_HEADER,
	VISIO_SWIMLANE_TITLE,
	type VisioSwimlaneDirection,
} from './edit-swimlane-commands';

export { VISIO_SEPARATOR_ROW };
export const VISIO_SWIMLANE_TITLE_ROW = 'ooxmlSwimlaneTitle';

/** Lane fill and outline per Style number (1 is Visio's warm default). */
export const SWIMLANE_STYLES: Record<number, { fill: string; line: string }> = {
	1: { fill: '#FFF9F1', line: '#C99A55' },
	2: { fill: '#F1F6FC', line: '#5B9BD5' },
	3: { fill: '#F5F5F5', line: '#7F7F7F' },
	4: { fill: '#F1F8F0', line: '#6AA84F' },
};
const look = (style: number) => SWIMLANE_STYLES[style] ?? SWIMLANE_STYLES[1]!;

/** Set (or add) a User row's Value cell. */
export function setUserRow(shape: Element, name: string, value: string, unit = 'STR'): void {
	let section = children(shape, 'Section').find((item) => attribute(item, 'N') === 'User');
	if (!section) {
		section = userSection(shape, []);
		const before = children(shape, 'Text')[0];
		shape.insertBefore(section, before ?? null);
	}
	let row = children(section, 'Row').find((item) => attribute(item, 'N') === name);
	if (!row) {
		row = node(shape, 'Row');
		row.setAttribute('N', name);
		section.appendChild(row);
	}
	let cell = children(row, 'Cell').find((item) => attribute(item, 'N') === 'Value');
	if (!cell) {
		valueCell(row, 'Value', value, unit);
		cell = children(row, 'Cell').find((item) => attribute(item, 'N') === 'Value')!;
	}
	cell.setAttribute('V', value);
	// An empty unit is Visio's plain number (numLanes, visCFFStyle and the other flowchart rows).
	if (unit) cell.setAttribute('U', unit);
	else cell.removeAttribute('U');
	cell.removeAttribute('F');
}
/** Set a colour or string cell, creating it when the sheet has none. */
export function setTextCell(shape: Element, name: string, value: string): void {
	const existing = cells(shape).get(name);
	if (existing) {
		existing.setAttribute('V', value);
		existing.removeAttribute('F');
	} else valueCell(shape, name, value);
}
const PI_HALF = Math.PI / 2;

/**
 * Where a lane's heading text sits: a band of VISIO_SWIMLANE_HEADER at the left of a horizontal
 * lane (text reading upwards) or at the top of a vertical one; `rotate` turns the text block
 * the other way (Rotate Lane Labels). Formulas keep the block with the lane's size.
 */
export function applyLaneLabel(
	lane: Element,
	direction: VisioSwimlaneDirection,
	rotate: boolean,
): void {
	const h = VISIO_SWIMLANE_HEADER;
	const width = numeric(cells(lane).get('Width'), 1);
	const height = numeric(cells(lane).get('Height'), 1);
	const horizontal = direction === 'horizontal';
	// Text reads along the band's long side unless rotated: vertical text in a horizontal lane.
	const turned = horizontal !== rotate;
	const along = horizontal ? height : width;
	const alongFormula = horizontal ? 'Height*1' : 'Width*1';
	const pinX = horizontal ? ([h / 2] as const) : ([width / 2, 'Width*0.5'] as const);
	const pinY = horizontal
		? ([height / 2, 'Height*0.5'] as const)
		: ([height - h / 2, `Height-${h / 2} in`] as const);
	setCell(lane, 'TxtPinX', pinX[0], pinX[1]);
	setCell(lane, 'TxtPinY', pinY[0], pinY[1]);
	// The text block's own width and height (before the angle turns it).
	const block: [number, string | undefined, number, string | undefined] = turned
		? [along, alongFormula, h, undefined]
		: horizontal
			? [h, undefined, along, alongFormula]
			: [along, alongFormula, h, undefined];
	setCell(lane, 'TxtWidth', block[0], block[1]);
	setCell(lane, 'TxtHeight', block[2], block[3]);
	setCell(lane, 'TxtLocPinX', block[0] / 2, 'TxtWidth*0.5');
	setCell(lane, 'TxtLocPinY', block[2] / 2, 'TxtHeight*0.5');
	setCell(lane, 'TxtAngle', turned ? PI_HALF : 0);
	setCell(lane, 'VerticalAlign', 1);
}

const box = (): [number, number][] => [
	[0, 0],
	[1, 0],
	[1, 1],
	[0, 1],
];
/** The heading band's inner edge as a relative line (the geometry editor admits only proportional rows). */
function dividerLine(direction: VisioSwimlaneDirection, box: VisioPartBox): [number, number][] {
	const h = VISIO_SWIMLANE_HEADER;
	return direction === 'horizontal'
		? [
				[Math.min(1, h / box.width), 0],
				[Math.min(1, h / box.width), 1],
			]
		: [
				[0, Math.max(0, 1 - h / box.height)],
				[1, Math.max(0, 1 - h / box.height)],
			];
}

const laneGeometry = (shape: Element, direction: VisioSwimlaneDirection, size: VisioPartBox) =>
	outlineGeometrySections(shape, { paths: [box()], lines: [dividerLine(direction, size)] });
/** Rebuild a lane's outline and heading divider (the direction changed); keeps the sections' place. */
export function applyLaneGeometry(
	lane: Element,
	direction: VisioSwimlaneDirection,
	size: VisioPartBox,
): void {
	const old = children(lane, 'Section').filter((item) => attribute(item, 'N') === 'Geometry');
	for (const section of laneGeometry(lane, direction, size))
		lane.insertBefore(section, old[0] ?? null);
	for (const section of old) lane.removeChild(section);
}

/** The flowchart rows Visio's categories and list machinery key on (see createSwimlanes). */
const EXCLUDED_FROM_LANES = 'CFF Container;CFF List;Swimlane;Phase;NonCFF';

/** A lane: a container with a heading band, no resize behaviour of its own (the frame stacks it). */
export function createLaneShape(
	root: Element,
	document: Element,
	id: string,
	frame: VisioPartBox,
	heading: string,
	direction: VisioSwimlaneDirection,
	rotate: boolean,
	style: number,
): Element {
	const shape = createShape(root, id);
	applyShapeCreationStyles(shape, document);
	placeBox(shape, frame);
	setCell(shape, 'LineWeight', 0.01);
	setCell(shape, 'LinePattern', 1);
	setCell(shape, 'FillPattern', 1);
	valueCell(shape, 'LineColor', look(style).line);
	valueCell(shape, 'FillForegnd', look(style).fill);
	applyLaneLabel(shape, direction, rotate);
	for (const section of laneGeometry(shape, direction, frame)) shape.appendChild(section);
	const paragraph = node(shape, 'Section');
	paragraph.setAttribute('N', 'Paragraph');
	const row = node(shape, 'Row');
	row.setAttribute('IX', '0');
	setCell(row, 'HorzAlign', 1);
	paragraph.appendChild(row);
	shape.appendChild(paragraph);
	shape.appendChild(
		userSection(shape, [
			['msvStructureType', 'Container'],
			['msvShapeCategories', 'Swimlane;DoNotContain'],
			['msvSDContainerExcludedCategories', EXCLUDED_FROM_LANES],
			['msvSDContainerNoRibbon', '1', 'BOOL'],
			[VISIO_CONTAINER_MARGIN_ROW, '0', 'IN'],
			[VISIO_CONTAINER_RESIZE_ROW, '0', 'STR'],
			['visCFFStyle', String(style), ''],
			[VISIO_SWIMLANE_LANE_ROW, '1', 'STR'],
		]),
	);
	shape.appendChild(textNode(shape, heading));
	return shape;
}

const swimlane_level = (kind: 'swimlane' | 'phase') => (kind === 'swimlane' ? -30080 : -30020);
/**
 * Visio's Swimlane List (members: the lanes in order) or Phase List (the separators): an
 * invisible 2D sheet spanning the lanes. `direction` is msvSDListDirection (0 left to right,
 * 2 top to bottom): lanes stack top to bottom in a horizontal flowchart, phases run left to right.
 */
export function createListShape(
	root: Element,
	document: Element,
	id: string,
	area: VisioPartBox,
	kind: 'swimlane' | 'phase',
	direction: 0 | 2,
	style: number,
): Element {
	const shape = createShape(root, id);
	applyShapeCreationStyles(shape, document);
	placeBox(shape, area);
	setCell(shape, 'LinePattern', 0);
	setCell(shape, 'FillPattern', 0);
	setCell(shape, 'ObjType', 4);
	// The cells of Visio's own list masters: no handles, glue to the edge only, drawn far behind.
	setCell(shape, 'NoObjHandles', 1);
	setCell(shape, 'GlueType', 8);
	setCell(shape, 'ShapePlaceStyle', 15);
	setCell(shape, 'IsTextEditTarget', 0);
	const level = swimlane_level(kind);
	setCell(shape, 'DisplayLevel', level, `GUARD(${level})`);
	for (const lock of ['LockRotate', 'LockGroup', 'LockCalcWH', 'LockReplace'])
		setCell(shape, lock, 1);
	for (const section of outlineGeometrySections(shape, { paths: [box()] }))
		shape.appendChild(section);
	const swimlane = kind === 'swimlane';
	const rows: [string, string, string?][] = [
		['msvStructureType', 'List'],
		['msvShapeCategories', 'CFF List;DoNotContain'],
		['msvSDListRequiredCategories', swimlane ? 'Swimlane' : 'Phase'],
		['msvSDContainerExcludedCategories', 'CFF Container;CFF List;Phase;NonCFF'],
		['msvSDContainerLocked', swimlane ? '0' : '1', 'BOOL'],
		['msvSDContainerMargin', '0'],
		['msvSDContainerNoHighlight', '1', 'BOOL'],
		['msvSDContainerNoRibbon', '1', 'BOOL'],
		['msvSDContainerResize', '0'],
		['msvSDContainerStyle', String(style), ''],
		['msvSDListSpacing', '0'],
		['msvSDListDirection', String(direction), ''],
		['msvSDListAlignment', '0'],
		['msvSDListItemMaster', '0'],
		['msvSDListNoReorder', '0'],
		['msvSDMembersOnHiddenLayer', '1', 'BOOL'],
		['visRotateLabel', '0', ''],
		['visVersion', '15', ''],
	];
	// Visio's Phase List only takes phases; the Swimlane List sets no such category.
	if (!swimlane) rows.splice(3, 0, ['msvSDContainerRequiredCategories', 'Phase']);
	shape.appendChild(userSection(shape, rows));
	return shape;
}

/** Restyle a lane or the frame outline for a Style number. */
export function applyLaneStyle(shape: Element, style: number, frame: boolean): void {
	setTextCell(shape, 'LineColor', look(style).line);
	if (!frame) setTextCell(shape, 'FillForegnd', look(style).fill);
}

/** The frame's title text block, anchored to the top and following the frame's height. */
export function applyFrameTitle(frame: Element, visible: boolean): void {
	const width = numeric(cells(frame).get('Width'), 1);
	const height = numeric(cells(frame).get('Height'), 1);
	const t = VISIO_SWIMLANE_TITLE;
	setCell(frame, 'TxtPinX', width / 2, 'Width*0.5');
	setCell(frame, 'TxtPinY', height - t / 2, `Height-${t / 2} in`);
	setCell(frame, 'TxtWidth', width, 'Width*1');
	setCell(frame, 'TxtHeight', visible ? t : 0.01);
	setCell(frame, 'TxtLocPinX', width / 2, 'TxtWidth*0.5');
	setCell(frame, 'TxtLocPinY', (visible ? t : 0.01) / 2, 'TxtHeight*0.5');
	setCell(frame, 'VerticalAlign', 1);
}

/** A phase separator: a thin bar across the lanes (horizontal flows get a vertical bar). */
export function createSeparatorShape(
	root: Element,
	document: Element,
	id: string,
	boxIn: VisioPartBox,
	style: number,
	visible: boolean,
	direction: VisioSwimlaneDirection = 'horizontal',
	label = '',
): Element {
	const shape = createShape(root, id);
	applyShapeCreationStyles(shape, document);
	placeBox(shape, boxIn);
	setCell(shape, 'LinePattern', 0);
	setCell(shape, 'FillPattern', visible ? 1 : 0);
	valueCell(shape, 'FillForegnd', look(style).line);
	for (const section of outlineGeometrySections(shape, { paths: [box()] }))
		shape.appendChild(section);
	const paragraph = node(shape, 'Section');
	paragraph.setAttribute('N', 'Paragraph');
	const row = node(shape, 'Row');
	row.setAttribute('IX', '0');
	setCell(row, 'HorzAlign', 0);
	paragraph.appendChild(row);
	shape.appendChild(paragraph);
	shape.appendChild(
		userSection(shape, [
			['msvStructureType', 'Container'],
			['msvShapeCategories', 'Phase;DoNotContain'],
			['msvSDContainerExcludedCategories', EXCLUDED_FROM_LANES],
			['msvSDContainerNoHighlight', '1', 'BOOL'],
			['msvSDContainerNoRibbon', '1', 'BOOL'],
			['visCFFStyle', String(style), ''],
			[VISIO_SEPARATOR_ROW, '1'],
		]),
	);
	shape.appendChild(textNode(shape, label));
	applySeparatorLabel(shape, direction, visible);
	return shape;
}

const PHASE_LABEL = { width: 1.4, height: 0.3 };
/**
 * Where a phase's name sits: beside the top of the bar in a horizontal flowchart (the bar is
 * vertical), above its left end in a vertical one. Hidden separators shrink the block away.
 */
export function applySeparatorLabel(
	bar: Element,
	direction: VisioSwimlaneDirection,
	visible: boolean,
): void {
	const thickness = numeric(
		cells(bar).get(direction === 'horizontal' ? 'Width' : 'Height'),
		VISIO_SEPARATOR_THICKNESS,
	);
	const w = visible ? PHASE_LABEL.width : 0.01;
	const h = visible ? PHASE_LABEL.height : 0.01;
	setCell(bar, 'TxtWidth', w);
	setCell(bar, 'TxtHeight', h);
	setCell(bar, 'TxtLocPinX', w / 2, 'TxtWidth*0.5');
	setCell(bar, 'TxtLocPinY', h / 2, 'TxtHeight*0.5');
	if (direction === 'horizontal') {
		setCell(bar, 'TxtPinX', thickness + w / 2);
		setCell(bar, 'TxtPinY', numeric(cells(bar).get('Height'), 1) - h / 2, `Height-${h / 2} in`);
	} else {
		setCell(bar, 'TxtPinX', w / 2);
		setCell(bar, 'TxtPinY', thickness + h / 2);
	}
	setCell(bar, 'VerticalAlign', 1);
}
