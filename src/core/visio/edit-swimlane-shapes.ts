import { attribute, children } from './sheet';
import { cells, numeric, setCell } from './edit-geometry-cells';
import {
	applyShapeCreationStyles,
	createShape,
	outlineGeometrySections,
} from './edit-shape-create';
import { node, placeBox, textNode, userSection, valueCell } from './edit-diagram-parts';
import { VISIO_SWIMLANE_LANE_ROW } from './shape-structure';
import {
	VISIO_CONTAINER_MARGIN_ROW,
	VISIO_CONTAINER_RESIZE_ROW,
	type VisioPartBox,
} from './edit-diagram-parts-commands';
import {
	VISIO_SWIMLANE_HEADER,
	VISIO_SWIMLANE_TITLE,
	type VisioSwimlaneDirection,
} from './edit-swimlane-commands';

export const VISIO_SEPARATOR_ROW = 'ooxmlSeparator';
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
	cell.setAttribute('U', unit);
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
	for (const section of outlineGeometrySections(shape, {
		paths: [box()],
		lines: [dividerLine(direction, frame)],
	}))
		shape.appendChild(section);
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
			[VISIO_CONTAINER_MARGIN_ROW, '0', 'IN'],
			[VISIO_CONTAINER_RESIZE_ROW, '0', 'STR'],
			[VISIO_SWIMLANE_LANE_ROW, '1', 'STR'],
		]),
	);
	shape.appendChild(textNode(shape, heading));
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
): Element {
	const shape = createShape(root, id);
	applyShapeCreationStyles(shape, document);
	placeBox(shape, boxIn);
	setCell(shape, 'LinePattern', 0);
	setCell(shape, 'FillPattern', visible ? 1 : 0);
	valueCell(shape, 'FillForegnd', look(style).line);
	for (const section of outlineGeometrySections(shape, { paths: [box()] }))
		shape.appendChild(section);
	shape.appendChild(userSection(shape, [[VISIO_SEPARATOR_ROW, '1']]));
	return shape;
}
