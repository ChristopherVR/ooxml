import { attribute } from './sheet';
import { fail } from './package-common';
import { cells, numeric, setCell } from './edit-geometry-cells';
import { createBox } from './edit-shape-create';
import { visioPathBounds } from './path-fit';
import type { VisioPathCreateEdit } from './edit-path-commands';

/** Relative arcs keep their axis ratio proportional to Width / Height so resizing scales them. */
const ARC_RATIO = /^WIDTH\/HEIGHT\*(\d+(?:\.\d+)?(?:E[+-]?\d+)?)$/;
const MINIMUM_EXTENT = 1e-6;
/** Plain decimal text: ShapeSheet formulas are written without exponent notation. */
const decimal = (value: number) =>
	/e/i.test(String(value)) ? value.toFixed(20).replace(/\.?0+$/, '') : String(value);

/** The frame a path creation occupies: centre pin and positive dimensions. */
export function visioPathFrame(edit: VisioPathCreateEdit): {
	minX: number;
	minY: number;
	width: number;
	height: number;
} {
	const bounds = visioPathBounds(edit, edit.segments);
	const width = bounds.maxX - bounds.minX,
		height = bounds.maxY - bounds.minY;
	if (!(width > MINIMUM_EXTENT) || !(height > MINIMUM_EXTENT))
		fail('INVALID_EDIT', 'A drawn path needs positive width and height; draw a line instead.');
	return { minX: bounds.minX, minY: bounds.minY, width, height };
}

/**
 * Freeform, Pencil and Arc output: a local 2D shape whose single Geometry section uses only
 * relative rows (RelMoveTo, RelLineTo, RelCubBezTo, RelEllipticalArcTo), so every coordinate is a
 * fraction of Width and Height and the shape moves, rotates and resizes without formula rewrites.
 * An open path is unfilled (NoFill 1); a closed one is filled.
 */
export function createPath(
	root: Element,
	document: Element,
	edit: VisioPathCreateEdit,
): { x: number; y: number; width: number; height: number } {
	const frame = visioPathFrame(edit);
	const expected = {
		x: frame.minX + frame.width / 2,
		y: frame.minY + frame.height / 2,
		width: frame.width,
		height: frame.height,
	};
	const shape = createBox(root, document, {
		type: 'create-rectangle',
		pageId: edit.pageId,
		shapeId: edit.shapeId,
		...expected,
	});
	const doc = root.ownerDocument!;
	const node = (name: string) => doc.createElementNS(root.namespaceURI, name);
	const fx = (x: number) => (x - frame.minX) / frame.width,
		fy = (y: number) => (y - frame.minY) / frame.height;
	const section = node('Section');
	section.setAttribute('N', 'Geometry');
	section.setAttribute('IX', '0');
	for (const [name, value] of [
		['NoFill', edit.closed ? 0 : 1],
		['NoLine', 0],
		['NoShow', 0],
		['NoSnap', 0],
		['NoQuickDrag', 0],
	] as const)
		setCell(section, name, value);
	let index = 0;
	const row = (type: string, values: [string, number, string?][]) => {
		const element = node('Row');
		element.setAttribute('IX', String(++index));
		element.setAttribute('T', type);
		for (const [name, value, formula] of values) setCell(element, name, value, formula);
		section.appendChild(element);
	};
	row('RelMoveTo', [
		['X', fx(edit.x)],
		['Y', fy(edit.y)],
	]);
	for (const segment of edit.segments) {
		const end: [string, number][] = [
			['X', fx(segment.x)],
			['Y', fy(segment.y)],
		];
		if (segment.kind === 'line') row('RelLineTo', end);
		else if (segment.kind === 'cubic')
			row('RelCubBezTo', [
				...end,
				['A', fx(segment.x1)],
				['B', fy(segment.y1)],
				['C', fx(segment.x2)],
				['D', fy(segment.y2)],
			]);
		else {
			const factor = (segment.ratio * frame.height) / frame.width;
			row('RelEllipticalArcTo', [
				...end,
				['A', fx(segment.a)],
				['B', fy(segment.b)],
				['C', 0],
				['D', segment.ratio, `Width/Height*${decimal(factor)}`],
			]);
		}
	}
	shape.appendChild(section);
	return expected;
}

/** Resize proof for a relative elliptical arc: no rotation and a Width/Height-proportional ratio. */
export function assertRelativeArcResizeRow(shape: Element, row: Element): void {
	const local = cells(shape),
		values = cells(row);
	const angle = values.get('C'),
		ratio = values.get('D');
	const match = attribute(ratio, 'F')?.replace(/\s/g, '').toUpperCase().match(ARC_RATIO);
	const expected = match
		? (numeric(local.get('Width')) / numeric(local.get('Height'))) * Number(match[1])
		: NaN;
	if (
		!angle ||
		attribute(angle, 'F') !== undefined ||
		numeric(angle) !== 0 ||
		!ratio ||
		!(Math.abs(numeric(ratio) - expected) <= 1e-9 * Math.max(1, Math.abs(expected)))
	)
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Relative elliptical arcs resize only with zero rotation and a Width/Height*k axis ratio.',
		);
}
