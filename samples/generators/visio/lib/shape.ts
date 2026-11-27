/**
 * One description of a 2-D shape, written three ways: as a master (formulas), as a shape drawn
 * straight on a page (formulas too) and as an instance of a master (cached overrides only).
 */
import { type Co, type Path, type Size, along, coFormula, geometry } from './geometry';
import type { StyleBook } from './styles';
import { attrs, cell, mergeCells, quoted, row, section } from './xml';

/** Shape Data types (the Type cell of a Property row). */
export const PROP = { string: 0, list: 1, number: 2, boolean: 3, date: 5, currency: 7 } as const;

export interface PropDef {
	/** Row name, read as `Prop.<name>`. */
	name: string;
	label: string;
	type: (typeof PROP)[keyof typeof PROP];
	/** A format picture, or the `;` separated choices of a list. */
	format?: string;
	prompt?: string;
	value?: string | number;
	/** An instance overriding the value of a row its master declares. */
	valueOnly?: boolean;
}

export interface UserDef {
	name: string;
	value: string | number;
	unit?: string;
	formula?: string;
}

/** Where the text block sits: its centre and size as fractions (plus inches) of the shape. */
export interface TextBlock {
	x: Co;
	y: Co;
	w: Co;
	h: Co;
}

export interface ShapeDef {
	/** One style name for line, fill and text, or one per kind. */
	style: string | { line: string; fill: string; text: string };
	paths: readonly Path[];
	textBlock?: TextBlock;
	/** Connection points as fractions of the shape; rows are glued to by index. */
	connections?: readonly (readonly [Co, Co])[];
	props?: readonly PropDef[];
	user?: readonly UserDef[];
	/** Extra ShapeSheet cells written as given. */
	cells?: readonly string[];
	/** Sub-shapes; their presence makes the shape a group. */
	children?: readonly ChildDef[];
}

export interface ChildDef extends ShapeDef {
	/** Left, bottom, width and height inside the parent. */
	box: { x: Co; y: Co; w: Co; h: Co };
}

/** The twelve points every box-like master offers: mid-sides first, then the quarter points. */
export const BOX_POINTS: readonly (readonly [number, number])[] = [
	[0.5, 1],
	[1, 0.5],
	[0.5, 0],
	[0, 0.5],
	[0.25, 1],
	[0.75, 1],
	[1, 0.75],
	[1, 0.25],
	[0.75, 0],
	[0.25, 0],
	[0, 0.25],
	[0, 0.75],
];
export const SIDE = { top: 0, right: 1, bottom: 2, left: 3 } as const;

export function styleAttrs(style: ShapeDef['style'], styles: StyleBook): string {
	const named = typeof style === 'string' ? { line: style, fill: style, text: style } : style;
	return attrs({
		LineStyle: styles.id(named.line),
		FillStyle: styles.id(named.fill),
		TextStyle: styles.id(named.text),
	});
}

function propValue(prop: PropDef): string {
	const value = prop.value ?? (prop.type === PROP.number || prop.type === PROP.currency ? 0 : '');
	if (prop.type === PROP.string || prop.type === PROP.list)
		return cell('Value', String(value), undefined, 'STR');
	if (prop.type === PROP.date) return cell('Value', String(value), undefined, 'DATE');
	if (prop.type === PROP.boolean) return cell('Value', value ? 1 : 0, undefined, 'BOOL');
	return cell('Value', value);
}

export function propertySection(props: readonly PropDef[] | undefined): string {
	return section(
		'Property',
		(props ?? []).map((prop, index) =>
			row({ N: prop.name }, [propValue(prop), ...(prop.valueOnly ? [] : rowCells(prop, index))]),
		),
	);
}

function rowCells(prop: PropDef, index: number): string[] {
	return [
		cell('Prompt', prop.prompt ?? ''),
		cell('Label', prop.label),
		cell('Format', prop.format ?? ''),
		cell('SortKey', String(index + 1)),
		cell('Type', prop.type),
		cell('Invisible', 0),
		cell('Verify', 0),
		cell('DataLinked', 0),
		cell('LangID', 'en-US'),
		cell('Calendar', 0),
	];
}

export function userSection(user: readonly UserDef[] | undefined): string {
	return section(
		'User',
		(user ?? []).map((entry) =>
			row({ N: entry.name }, [
				typeof entry.value === 'number'
					? cell('Value', entry.value, entry.formula, entry.unit)
					: cell('Value', entry.value, entry.formula ?? quoted(entry.value), entry.unit ?? 'STR'),
				cell('Prompt', ''),
			]),
		),
	);
}

type Mode = 'formula' | 'inherited';

function connectionSection(def: ShapeDef, size: Size, mode: Mode): string {
	return section(
		'Connection',
		(def.connections ?? []).map(([x, y], index) =>
			row({ T: 'Connection', IX: index }, [
				cell('X', along(x, size.w), mode === 'inherited' ? 'Inh' : coFormula(x, 'Width')),
				cell('Y', along(y, size.h), mode === 'inherited' ? 'Inh' : coFormula(y, 'Height')),
				...(mode === 'inherited'
					? []
					: [cell('DirX', 0), cell('DirY', 0), cell('Type', 0), cell('AutoGen', 0)]),
			]),
		),
	);
}

function textBlockCells(block: TextBlock | undefined, size: Size, mode: Mode): string[] {
	if (!block) return [];
	const f = (text: string) => (mode === 'inherited' ? 'Inh' : text);
	const w = along(block.w, size.w);
	const h = along(block.h, size.h);
	return [
		cell('TxtPinX', along(block.x, size.w), f(coFormula(block.x, 'Width'))),
		cell('TxtPinY', along(block.y, size.h), f(coFormula(block.y, 'Height'))),
		cell('TxtWidth', w, f(coFormula(block.w, 'Width'))),
		cell('TxtHeight', h, f(coFormula(block.h, 'Height'))),
		cell('TxtLocPinX', w / 2, f('TxtWidth*0.5')),
		cell('TxtLocPinY', h / 2, f('TxtHeight*0.5')),
		...(mode === 'inherited' ? [] : [cell('TxtAngle', 0)]),
	];
}

/** Size and position of a child inside a parent of the given size. */
function childBox(child: ChildDef, parent: Size): { x: number; y: number; size: Size } {
	const size = { w: along(child.box.w, parent.w), h: along(child.box.h, parent.h) };
	return {
		x: along(child.box.x, parent.w) + size.w / 2,
		y: along(child.box.y, parent.h) + size.h / 2,
		size,
	};
}

export interface Placement {
	x: number;
	y: number;
	size: Size;
}

/** Counts the shapes of a definition: itself and every descendant. */
export function shapeCount(def: ShapeDef): number {
	return 1 + (def.children ?? []).reduce((total, child) => total + shapeCount(child), 0);
}

/**
 * A complete shape: transform, formatting, sections and geometry with formulas. `ids` hands out
 * the shape IDs (a master numbers its shapes from 5, a page from its own counter). `parent`
 * is the sheet a child measures itself against.
 */
export function drawnShape(
	def: ShapeDef,
	place: Placement,
	styles: StyleBook,
	ids: () => number,
	extra: {
		head?: string;
		body?: string;
		parent?: { id: number; child: ChildDef };
		/** A shape on a page takes its theme colours as local theme formulas. */
		onPage?: boolean;
	} = {},
): string {
	const id = ids();
	const { size } = place;
	const pin = extra.parent
		? (() => {
				const { box } = extra.parent.child;
				const sheet = `Sheet.${extra.parent.id}!`;
				const w = coFormula(box.w, `${sheet}Width`);
				const h = coFormula(box.h, `${sheet}Height`);
				return [
					cell('PinX', place.x, `${coFormula(box.x, `${sheet}Width`)}+(${w})*0.5`),
					cell('PinY', place.y, `${coFormula(box.y, `${sheet}Height`)}+(${h})*0.5`),
					cell('Width', size.w, w),
					cell('Height', size.h, h),
				];
			})()
		: [cell('PinX', place.x), cell('PinY', place.y), cell('Width', size.w), cell('Height', size.h)];
	const cells = mergeCells(
		pin,
		[
			cell('LocPinX', size.w / 2, 'Width*0.5'),
			cell('LocPinY', size.h / 2, 'Height*0.5'),
			cell('Angle', 0),
			cell('FlipX', 0),
			cell('FlipY', 0),
			cell('ResizeMode', 0),
		],
		textBlockCells(def.textBlock, size, 'formula'),
		// A group draws its own geometry behind its members (the default is in front).
		def.children?.length ? [cell('DisplayMode', 1)] : [],
		extra.onPage ? styles.themeCells(def.style) : [],
		def.cells,
	);
	const children = (def.children ?? []).map((child) =>
		drawnShape(child, childBox(child, size), styles, ids, {
			parent: { id, child },
			...(extra.onPage ? { onPage: true } : {}),
		}),
	);
	return (
		`<Shape${attrs({ ID: id, Type: children.length ? 'Group' : 'Shape' })}${extra.head ?? ''}${styleAttrs(def.style, styles)}>` +
		cells.join('') +
		connectionSection(def, size, 'formula') +
		propertySection(def.props) +
		userSection(def.user) +
		geometry(def.paths, size) +
		(extra.body ?? '') +
		(children.length ? `<Shapes>${children.join('')}</Shapes>` : '') +
		'</Shape>'
	);
}

/**
 * An instance of a master: the pin, and for a size other than the master's the cached values of
 * everything that follows the size (the form Visio saves, `F="Inh"`), so a reader that does not
 * evaluate formulas still draws it right.
 */
export function instanceShape(
	def: ShapeDef,
	masterSize: Size,
	place: Placement,
	ids: () => number,
	masterShapeIds: () => number,
	extra: {
		head?: string;
		cells?: readonly string[];
		body?: string;
		top?: boolean;
		/** Gives each shape of the instance its theme colours as local theme formulas. */
		styles?: StyleBook;
		/** Local cells for sub-shapes, by their shape ID in the master. */
		childCells?: Readonly<Record<number, readonly string[]>>;
	} = {},
): string {
	const id = ids();
	const masterShapeId = masterShapeIds();
	const { size } = place;
	const resized = Math.abs(size.w - masterSize.w) > 1e-9 || Math.abs(size.h - masterSize.h) > 1e-9;
	const top = extra.top !== false;
	const cells: string[] = [];
	if (top) {
		cells.push(cell('PinX', place.x), cell('PinY', place.y));
		cells.push(
			cell('Width', size.w, resized ? undefined : 'Inh'),
			cell('Height', size.h, resized ? undefined : 'Inh'),
		);
	} else if (resized) {
		cells.push(
			cell('PinX', place.x, 'Inh'),
			cell('PinY', place.y, 'Inh'),
			cell('Width', size.w, 'Inh'),
			cell('Height', size.h, 'Inh'),
		);
	}
	if (top || resized)
		cells.push(cell('LocPinX', size.w / 2, 'Inh'), cell('LocPinY', size.h / 2, 'Inh'));
	if (resized) cells.push(...textBlockCells(def.textBlock, size, 'inherited'));
	const local = mergeCells(
		cells,
		extra.styles?.themeCells(def.style),
		extra.cells,
		extra.childCells?.[masterShapeId],
	);
	const children = (def.children ?? []).map((child) => {
		const box = childBox(child, size);
		const masterBox = childBox(child, masterSize);
		return instanceShape(child, masterBox.size, box, ids, masterShapeIds, {
			top: false,
			...(extra.styles ? { styles: extra.styles } : {}),
			...(extra.childCells ? { childCells: extra.childCells } : {}),
		});
	});
	const identity = top ? '' : attrs({ MasterShape: masterShapeId });
	return (
		`<Shape${attrs({ ID: id, Type: children.length ? 'Group' : 'Shape' })}${extra.head ?? ''}${identity}>` +
		local.join('') +
		(resized ? connectionSection(def, size, 'inherited') : '') +
		(resized ? geometry(def.paths, size, 'inherited') : '') +
		(extra.body ?? '') +
		(children.length ? `<Shapes>${children.join('')}</Shapes>` : '') +
		'</Shape>'
	);
}
