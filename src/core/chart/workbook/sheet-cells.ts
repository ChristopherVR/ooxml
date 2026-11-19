/**
 * Cell-level read and write helpers for a parsed SpreadsheetML worksheet part
 * (`xl/worksheets/sheetN.xml`), used by the embedded-workbook write-back path (`writer.ts`).
 *
 * Every mutation touches only the `<c>` elements a write targets. Everything else on the
 * worksheet (styles, formatting, other cells, other rows, extensions) is left as parsed, so an
 * untouched worksheet is never re-serialized. New rows and cells take the namespace and prefix
 * of the `sheetData` they are added to.
 *
 * @module chart/workbook/sheet-cells
 */
import { buildXml, elements, parseXml, type XmlDocument, type XmlElement } from '../../xml/index';
import { parseCellAddress } from './cellref';

/** One resolved cell write: a concrete "A1"-style ref plus its new value. */
export interface ChartCellWrite {
	ref: string;
	isNumeric: boolean;
	value: string;
}

/** Child elements of `parent` with local name `name`, in document order. */
export const childrenNamed = (parent: XmlElement | undefined, name: string): XmlElement[] =>
	parent ? elements(parent).filter((child) => child.localName === name) : [];

/** The first child element of `parent` with local name `name`. */
export const childNamed = (parent: XmlElement | undefined, name: string): XmlElement | undefined =>
	parent ? elements(parent).find((child) => child.localName === name) : undefined;

/** The `sheetData` element of a worksheet document, if it has one. */
export function worksheetData(doc: XmlDocument): XmlElement | undefined {
	const root = doc.documentElement;
	return root?.localName === 'worksheet' ? childNamed(root, 'sheetData') : undefined;
}

function create(near: XmlElement, name: string): XmlElement {
	const qualified = near.prefix ? `${near.prefix}:${name}` : name;
	return near.ownerDocument!.createElementNS(near.namespaceURI, qualified);
}

/** The rows and cells of a worksheet by number and reference, kept current as cells are added. */
interface SheetIndex {
	sheetData: XmlElement;
	rows: Map<number, XmlElement>;
	/** Row numbers in ascending order. */
	order: number[];
	cells: Map<string, XmlElement>;
}

function indexSheet(sheetData: XmlElement): SheetIndex {
	const rows = new Map<number, XmlElement>();
	const cells = new Map<string, XmlElement>();
	for (const row of childrenNamed(sheetData, 'row')) {
		const number = Number(row.getAttribute('r'));
		if (!rows.has(number)) rows.set(number, row);
		for (const cell of childrenNamed(row, 'c')) {
			const ref = cell.getAttribute('r');
			if (ref && !cells.has(ref)) cells.set(ref, cell);
		}
	}
	return { sheetData, rows, order: [...rows.keys()].sort((a, b) => a - b), cells };
}

/** The position of the first row numbered above `number`. */
function rowAfter(order: readonly number[], number: number): number {
	let low = 0;
	let high = order.length;
	while (low < high) {
		const middle = (low + high) >> 1;
		if (order[middle]! > number) high = middle;
		else low = middle + 1;
	}
	return low;
}

/** Get (creating if necessary) the row and cell for `ref`, each inserted in row or column order. */
function ensureCell(sheet: SheetIndex, ref: string): XmlElement | undefined {
	const existing = sheet.cells.get(ref);
	if (existing) return existing;
	const address = parseCellAddress(ref);
	if (!address) return undefined;
	const rowNumber = address.row + 1;
	let row = sheet.rows.get(rowNumber);
	if (!row) {
		row = create(sheet.sheetData, 'row');
		row.setAttribute('r', String(rowNumber));
		const at = rowAfter(sheet.order, rowNumber);
		const next = sheet.order[at];
		sheet.sheetData.insertBefore(row, next === undefined ? null : sheet.rows.get(next)!);
		sheet.order.splice(at, 0, rowNumber);
		sheet.rows.set(rowNumber, row);
	}
	const cell = create(row, 'c');
	cell.setAttribute('r', ref);
	const after = childrenNamed(row, 'c').find((candidate) => {
		const other = parseCellAddress(candidate.getAttribute('r') ?? '');
		return other !== undefined && other.col > address.col;
	});
	row.insertBefore(cell, after ?? null);
	// `spans` is an optional hint of the row's column range; a new cell can fall outside it.
	row.removeAttribute('spans');
	sheet.cells.set(ref, cell);
	return cell;
}

/** The text of an `inlineStr` cell (`<is><t>...</t></is>`), if any. */
function inlineText(cell: XmlElement): string | undefined {
	const text = childNamed(childNamed(cell, 'is'), 't');
	return text ? (text.textContent ?? '') : undefined;
}

/**
 * Whether `cell` already holds `value` in the representation a write would produce, so an
 * unchanged cell is left untouched rather than reformatted.
 */
function holds(
	cell: XmlElement,
	isNumeric: boolean,
	value: string,
	sharedStrings: readonly string[],
): boolean {
	const type = cell.getAttribute('t') || 'n';
	const current = childNamed(cell, 'v')?.textContent ?? undefined;
	if (value === '') return current === undefined && !childNamed(cell, 'is');
	if (isNumeric) {
		if (isCellError(value)) return type === 'e' && current === value;
		if (type !== 'n' || current === undefined) return false;
		const before = Number.parseFloat(current);
		const after = Number.parseFloat(value);
		return Number.isFinite(before) && Number.isFinite(after) ? before === after : current === value;
	}
	if (type === 's') return sharedStrings[Number(childNamed(cell, 'v')?.textContent)] === value;
	if (type === 'str') return childNamed(cell, 'v')?.textContent === value;
	return type === 'inlineStr' && inlineText(cell) === value;
}

/** A worksheet error (`#N/A`), which a numeric chart cache holds as its text. */
const isCellError = (value: string) => value.startsWith('#');

/** Overwrite a cell's type and value to hold `value`, dropping any prior value shape. */
function writeCell(cell: XmlElement, isNumeric: boolean, value: string): void {
	for (const child of elements(cell))
		if (child.localName === 'is' || child.localName === 'v' || child.localName === 'f')
			cell.removeChild(child);
	cell.removeAttribute('t');
	// A blank point empties the cell, as Excel stores a cleared cell.
	if (value === '') return;
	if (isNumeric) {
		if (isCellError(value)) cell.setAttribute('t', 'e');
		const v = create(cell, 'v');
		v.textContent = value;
		cell.appendChild(v);
		return;
	}
	// `inlineStr` keeps the write self-contained: it never touches the shared string table, so
	// other cells referencing the same shared string are unaffected. `xml:space="preserve"` keeps
	// boundary whitespace in an edited label.
	cell.setAttribute('t', 'inlineStr');
	const is = create(cell, 'is');
	const t = create(is, 't');
	t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
	t.textContent = value;
	is.appendChild(t);
	cell.appendChild(is);
}

/** A cell a write changed, as markup: before the write (absent when it made the cell) and after. */
export interface ChartCellChange {
	ref: string;
	before?: string;
	after: string;
}

/**
 * Apply a batch of cell writes to a parsed worksheet and return the cells it changed, so the
 * caller only re-serializes a worksheet that differs from what was loaded. `sharedStrings` (the
 * workbook's shared string texts) lets a shared-string cell that already shows the text stay.
 */
export function writeWorksheetCells(
	doc: XmlDocument,
	writes: readonly ChartCellWrite[],
	sharedStrings: readonly string[] = [],
): ChartCellChange[] {
	const sheetData = worksheetData(doc);
	if (!sheetData) return [];
	const sheet = indexSheet(sheetData);
	const changes = new Map<string, { before?: string; cell: XmlElement }>();
	for (const write of writes) {
		if (!parseCellAddress(write.ref)) continue;
		const existing = sheet.cells.get(write.ref);
		if (existing && holds(existing, write.isNumeric, write.value, sharedStrings)) continue;
		const before = existing && !changes.has(write.ref) ? buildXml(existing) : undefined;
		const cell = existing ?? ensureCell(sheet, write.ref);
		if (!cell) continue;
		writeCell(cell, write.isNumeric, write.value);
		if (!changes.has(write.ref))
			changes.set(write.ref, before === undefined ? { cell } : { before, cell });
	}
	return [...changes].map(([ref, { before, cell }]) => ({
		ref,
		...(before === undefined ? {} : { before }),
		after: buildXml(cell),
	}));
}

/**
 * Apply a batch of cell writes to a parsed worksheet. Returns whether any cell changed, so the
 * caller only re-serializes a worksheet that differs from what was loaded.
 */
export const applyCellWritesToWorksheet = (
	doc: XmlDocument,
	writes: readonly ChartCellWrite[],
	sharedStrings: readonly string[] = [],
): boolean => writeWorksheetCells(doc, writes, sharedStrings).length > 0;

const XMLNS = 'http://www.w3.org/2000/xmlns/';

/**
 * Puts back each cell of `changes` that still holds what its write left, so cells written since,
 * by another chart or by hand, stay. Returns the references it restored.
 */
export function restoreWorksheetCells(
	doc: XmlDocument,
	changes: readonly ChartCellChange[],
): string[] {
	const sheetData = worksheetData(doc);
	if (!sheetData) return [];
	const sheet = indexSheet(sheetData);
	const restored: string[] = [];
	for (const change of changes) {
		const cell = sheet.cells.get(change.ref);
		if (!cell || buildXml(cell) !== change.after) continue;
		if (change.before === undefined) {
			cell.parentNode?.removeChild(cell);
		} else {
			const before = parseXml(change.before, { label: 'Worksheet cell' }).documentElement!;
			for (const attribute of Array.from(cell.attributes)) cell.removeAttributeNode(attribute);
			for (const child of Array.from(cell.childNodes)) cell.removeChild(child);
			for (const attribute of Array.from(before.attributes))
				if (attribute.namespaceURI !== XMLNS)
					cell.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
			for (const child of Array.from(before.childNodes))
				cell.appendChild(doc.importNode(child, true));
		}
		restored.push(change.ref);
	}
	return restored;
}
