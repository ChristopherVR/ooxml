import { elements } from '../xml/index';
import { executableCellFormula } from './cell-formula';
import { assertShapeLocks } from './edit-style-admission';
import { fail } from './package-common';
import { decodeVisioPlainText, encodeVisioPlainText } from './plain-text';
import { attribute, child, children } from './sheet';

/**
 * What Edit Master changes on a master's shape: its default size (geometry), its fill and line
 * and the formatting of its text, besides its default text. Colours are `#rrggbb`, sizes inches,
 * the line weight and font size points.
 */
export interface VisioMasterFormat {
	text?: string;
	width?: number;
	height?: number;
	fillColor?: string;
	lineColor?: string;
	lineWeight?: number;
	fontColor?: string;
	fontSize?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	horizontalAlign?: 'left' | 'center' | 'right' | 'justify';
}

/** One cell Edit Master wrote: a shape cell, or a cell of a Character or Paragraph row. */
export interface VisioMasterCellChange {
	section?: 'Character' | 'Paragraph';
	/** The row IX of a section cell. */
	row?: string;
	name: string;
	value: string;
	unit?: string;
}

const COLOR = /^#[0-9a-f]{6}$/i;
const ALIGN = ['left', 'center', 'right', 'justify'] as const;

/** Copy and validate a master format; nothing is changed when every field is left out. */
export function snapshotVisioMasterFormat(input: VisioMasterFormat): VisioMasterFormat {
	if (typeof input !== 'object' || input === null) fail('INVALID_EDIT', 'Invalid master format.');
	const result: VisioMasterFormat = {};
	const size = (value: unknown, what: string, min: number, max: number) => {
		if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max)
			fail('INVALID_EDIT', `${what} must be between ${min} and ${max}.`);
		return value as number;
	};
	const color = (value: unknown) => {
		if (typeof value !== 'string' || !COLOR.test(value))
			fail('INVALID_EDIT', 'Colours must be #rrggbb.');
		return (value as string).toLowerCase();
	};
	if (input.text !== undefined) {
		if (
			typeof input.text !== 'string' ||
			input.text.length > 4096 ||
			/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/u.test(input.text)
		)
			fail('UNSUPPORTED_TEXT_EDIT', 'Enter plain text of up to 4096 characters.');
		result.text = input.text.replace(/\r\n?/g, '\n');
	}
	if (input.width !== undefined) result.width = size(input.width, 'Width', 0.01, 1000);
	if (input.height !== undefined) result.height = size(input.height, 'Height', 0.01, 1000);
	if (input.fillColor !== undefined) result.fillColor = color(input.fillColor);
	if (input.lineColor !== undefined) result.lineColor = color(input.lineColor);
	if (input.lineWeight !== undefined)
		result.lineWeight = size(input.lineWeight, 'Line weight', 0, 72);
	if (input.fontColor !== undefined) result.fontColor = color(input.fontColor);
	if (input.fontSize !== undefined) result.fontSize = size(input.fontSize, 'Font size', 1, 1638);
	for (const key of ['bold', 'italic', 'underline'] as const)
		if (input[key] !== undefined) {
			if (typeof input[key] !== 'boolean') fail('INVALID_EDIT', `${key} must be true or false.`);
			result[key] = input[key];
		}
	if (input.horizontalAlign !== undefined) {
		if (!ALIGN.includes(input.horizontalAlign))
			fail('INVALID_EDIT', 'Alignment must be left, center, right or justify.');
		result.horizontalAlign = input.horizontalAlign;
	}
	return result;
}

/** True when the format changes the master's size, which instances that inherit it follow. */
export const changesMasterSize = (format: VisioMasterFormat) =>
	format.width !== undefined || format.height !== undefined;

function cellOf(owner: Element, name: string): Element {
	const existing = children(owner, 'Cell').find((cell) => attribute(cell, 'N') === name);
	if (existing) return existing;
	const cell = owner.ownerDocument!.createElementNS(owner.namespaceURI, 'Cell');
	cell.setAttribute('N', name);
	const before = Array.from(owner.childNodes).find(
		(node) => node.nodeType === 1 && (node as Element).localName !== 'Cell',
	);
	owner.insertBefore(cell, before ?? null);
	return cell;
}

/** The rows of a Character or Paragraph section, row 0 first; created when the shape has none. */
function textRows(shape: Element, name: 'Character' | 'Paragraph'): Element[] {
	let section = children(shape, 'Section').find((node) => attribute(node, 'N') === name);
	if (!section) {
		section = shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Section');
		section.setAttribute('N', name);
		const before = children(shape, 'Section')[0] ?? child(shape, 'Text') ?? null;
		shape.insertBefore(section, before);
	}
	const rows = children(section, 'Row');
	if (!rows.some((row) => (attribute(row, 'IX') ?? '0') === '0')) {
		const row = shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Row');
		row.setAttribute('IX', '0');
		section.insertBefore(row, rows[0] ?? null);
		rows.unshift(row);
	}
	return rows.filter((row) => attribute(row, 'Del') !== '1');
}

const SIZE_RATIO = /^\s*(Width|Height)\s*\*\s*([-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?)\s*$/i;
const SIZE_REFERENCE = /(?<![.!\w])(Width|Height)(?![\w(])/i;

/**
 * Resize a master's shape, as editing Width and Height in its ShapeSheet does: cells computed as
 * a ratio of the size (`Width*0.5`, the usual pin, text block, control and absolute geometry
 * form) are recalculated; relative geometry follows by itself. A cell that computes from the
 * size in any other way is refused rather than left stale.
 */
function resize(shape: Element, format: VisioMasterFormat): VisioMasterCellChange[] {
	const changes: VisioMasterCellChange[] = [];
	const current = (name: string) =>
		Number(
			attribute(
				children(shape, 'Cell').find((cell) => attribute(cell, 'N') === name),
				'V',
			),
		);
	const size = {
		width: format.width ?? current('Width'),
		height: format.height ?? current('Height'),
	};
	if (!Number.isFinite(size.width) || !Number.isFinite(size.height))
		fail('UNSUPPORTED_MASTER_EDIT', 'The master has no size to change.');
	for (const cell of Array.from(shape.getElementsByTagNameNS(shape.namespaceURI, 'Cell'))) {
		const name = attribute(cell, 'N') ?? '';
		if (name === 'Width' || name === 'Height') continue;
		const formula = executableCellFormula(attribute(cell, 'F'));
		if (!formula || !SIZE_REFERENCE.test(formula)) continue;
		const ratio = SIZE_RATIO.exec(formula);
		if (!ratio)
			fail(
				'UNSUPPORTED_MASTER_EDIT',
				`The master computes ${name} from its size (${formula}), which Edit Master cannot update.`,
			);
		const dimension = ratio![1]!.toLowerCase() === 'width' ? size.width : size.height;
		cell.setAttribute('V', String(dimension * Number(ratio![2])));
	}
	for (const [name, value] of [
		['Width', format.width],
		['Height', format.height],
	] as const) {
		if (value === undefined) continue;
		const cell = cellOf(shape, name);
		cell.setAttribute('V', String(value));
		cell.setAttribute('U', 'IN');
		cell.removeAttribute('F');
		changes.push({ name, value: String(value), unit: 'IN' });
		// The master's pin stays at the middle of its master page, as Visio keeps it.
		const pin = children(shape, 'Cell').find((node) => attribute(node, 'N') === `Pin${name[0]}`);
		if (pin && !executableCellFormula(attribute(pin, 'F')))
			pin.setAttribute('V', String(value / 2));
	}
	return changes;
}

/** Set the master's default plain text; refused for rich text and fields. */
function setText(shape: Element, text: string): boolean {
	const existing = child(shape, 'Text');
	if (
		children(shape, 'Text').length > 1 ||
		(existing && elements(existing).length) ||
		children(shape, 'Section').some((section) => attribute(section, 'N') === 'Field')
	)
		fail('UNSUPPORTED_TEXT_EDIT', 'Select a top-level shape with editable plain text.');
	if (decodeVisioPlainText(existing?.textContent ?? '') === text) return false;
	const node = existing ?? shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Text');
	node.textContent = encodeVisioPlainText(text);
	if (!existing) shape.appendChild(node);
	return true;
}

/**
 * Apply Edit Master to the one top-level shape of a master: default text, size, fill, line and
 * text formatting are written as plain local values of the master (a themed or styled value
 * becomes the chosen one, as formatting a shape in Visio does). Locked cells are refused.
 * Returns the cells written, which inheriting instances then show, and whether anything changed.
 */
export function applyVisioMasterFormat(
	shape: Element,
	document: Element,
	format: VisioMasterFormat,
): { changes: VisioMasterCellChange[]; changed: boolean } {
	if (shape.hasAttribute('Master') || attribute(shape, 'Type') !== 'Shape')
		fail('UNSUPPORTED_MASTER_EDIT', 'Edit Master changes a master with one top-level shape.');
	const locks = [
		...(format.text !== undefined ? ['LockTextEdit'] : []),
		...(changesMasterSize(format) ? ['LockWidth', 'LockHeight'] : []),
		...(Object.keys(format).some((key) => key !== 'text' && key !== 'width' && key !== 'height')
			? ['LockFormat']
			: []),
	];
	if (locks.length) assertShapeLocks(shape, document, locks);
	let changed = format.text !== undefined && setText(shape, format.text);
	const changes: VisioMasterCellChange[] = changesMasterSize(format) ? resize(shape, format) : [];
	const write = (
		owner: Element,
		name: string,
		value: string,
		options: { unit?: string; formula?: string; section?: 'Character' | 'Paragraph' } = {},
	) => {
		const cell = cellOf(owner, name);
		cell.setAttribute('V', value);
		if (options.unit) cell.setAttribute('U', options.unit);
		else cell.removeAttribute('U');
		if (options.formula) cell.setAttribute('F', options.formula);
		else cell.removeAttribute('F');
		changes.push({
			name,
			value,
			...(options.unit ? { unit: options.unit } : {}),
			...(options.section ? { section: options.section, row: attribute(owner, 'IX') ?? '0' } : {}),
		});
	};
	const rgb = (hex: string) =>
		`RGB(${[1, 3, 5].map((index) => parseInt(hex.slice(index, index + 2), 16)).join(',')})`;
	if (format.fillColor !== undefined) {
		write(shape, 'FillPattern', '1');
		write(shape, 'FillGradientEnabled', '0');
		write(shape, 'FillForegnd', format.fillColor, { formula: rgb(format.fillColor) });
		write(shape, 'FillForegndTrans', '0');
	}
	if (format.lineColor !== undefined) {
		write(shape, 'LineColor', format.lineColor, { formula: rgb(format.lineColor) });
		write(shape, 'LineColorTrans', '0');
		write(shape, 'LineGradientEnabled', '0');
	}
	if (format.lineWeight !== undefined)
		write(shape, 'LineWeight', String(format.lineWeight / 72), { unit: 'PT' });
	const characterEdit =
		format.fontColor !== undefined ||
		format.fontSize !== undefined ||
		[format.bold, format.italic, format.underline].some((value) => value !== undefined);
	if (characterEdit)
		for (const row of textRows(shape, 'Character')) {
			const section = { section: 'Character' as const };
			if (format.fontSize !== undefined)
				write(row, 'Size', String(format.fontSize / 72), { ...section, unit: 'PT' });
			if (format.fontColor !== undefined) {
				write(row, 'Color', format.fontColor, { ...section, formula: rgb(format.fontColor) });
				write(row, 'ColorTrans', '0', section);
			}
			if ([format.bold, format.italic, format.underline].some((value) => value !== undefined)) {
				const style = children(row, 'Cell').find((cell) => attribute(cell, 'N') === 'Style');
				let bits = Number(attribute(style, 'V') ?? 0);
				if (!Number.isSafeInteger(bits) || bits < 0 || bits > 255) bits = 0;
				for (const [value, bit] of [
					[format.bold, 1],
					[format.italic, 2],
					[format.underline, 4],
				] as const)
					if (value !== undefined) bits = value ? bits | bit : bits & ~bit;
				write(row, 'Style', String(bits), section);
			}
		}
	if (format.horizontalAlign !== undefined)
		for (const row of textRows(shape, 'Paragraph'))
			write(row, 'HorzAlign', String(ALIGN.indexOf(format.horizontalAlign)), {
				section: 'Paragraph',
			});
	return { changes, changed: changed || changes.length > 0 };
}

/**
 * Show a master edit on one of its instances, as Visio does: a cell the instance inherits keeps
 * a cached copy of the master's value (`F="Inh"`), which is refreshed; a cell it sets itself is
 * the instance's own formatting and stays. Returns whether the instance changed.
 */
export function refreshInheritedCaches(
	instance: Element,
	changes: readonly VisioMasterCellChange[],
): boolean {
	let changed = false;
	for (const change of changes) {
		let owner: Element | undefined = instance;
		if (change.section) {
			const section = children(instance, 'Section').find(
				(node) => attribute(node, 'N') === change.section,
			);
			owner = children(section, 'Row').find((row) => (attribute(row, 'IX') ?? '0') === change.row);
		}
		const cell =
			owner && children(owner, 'Cell').find((node) => attribute(node, 'N') === change.name);
		if (!cell || attribute(cell, 'F') !== 'Inh') continue;
		if (attribute(cell, 'V') === change.value && attribute(cell, 'U') === change.unit) continue;
		cell.setAttribute('V', change.value);
		if (change.unit) cell.setAttribute('U', change.unit);
		else cell.removeAttribute('U');
		changed = true;
	}
	return changed;
}

/** The master shape's default plain text, empty when it has none. */
export const masterPlainText = (shape: Element): string =>
	decodeVisioPlainText(child(shape, 'Text')?.textContent ?? '');

/**
 * Let an instance show a master's new default text when its own text is only the empty
 * placeholder a drop from a master without text writes (`<Text/>`, so character formats have
 * somewhere to go). While the master had no text the two read the same; text the user typed, or
 * any text with fields or format markers, stays. Returns whether the instance changed.
 */
export function releasePlaceholderText(instance: Element): boolean {
	const texts = children(instance, 'Text');
	const text = texts[0];
	if (texts.length !== 1 || !text || elements(text).length || text.textContent) return false;
	instance.removeChild(text);
	return true;
}

/** True when the instance takes the master's size: it stores no Width or Height of its own. */
export function inheritsMasterSize(instance: Element): boolean {
	return ['Width', 'Height'].some((name) => {
		const cell = children(instance, 'Cell').find((node) => attribute(node, 'N') === name);
		return !cell || attribute(cell, 'F') === 'Inh';
	});
}
