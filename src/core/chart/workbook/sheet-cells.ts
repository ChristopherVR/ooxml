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
import { elements, type XmlDocument, type XmlElement } from '../../xml/index';
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

function findCell(sheetData: XmlElement, ref: string): XmlElement | undefined {
	for (const row of childrenNamed(sheetData, 'row'))
		for (const cell of childrenNamed(row, 'c')) if (cell.getAttribute('r') === ref) return cell;
	return undefined;
}

/** Get (creating if necessary) the row and cell for `ref`, each inserted in row or column order. */
function ensureCell(sheetData: XmlElement, ref: string): XmlElement | undefined {
	const address = parseCellAddress(ref);
	if (!address) return undefined;
	const rowNumber = address.row + 1;
	const rows = childrenNamed(sheetData, 'row');
	let row = rows.find((candidate) => Number(candidate.getAttribute('r')) === rowNumber);
	if (!row) {
		row = create(sheetData, 'row');
		row.setAttribute('r', String(rowNumber));
		const after = rows.find((candidate) => Number(candidate.getAttribute('r')) > rowNumber);
		sheetData.insertBefore(row, after ?? null);
	}
	const cells = childrenNamed(row, 'c');
	let cell = cells.find((candidate) => candidate.getAttribute('r') === ref);
	if (!cell) {
		cell = create(row, 'c');
		cell.setAttribute('r', ref);
		const after = cells.find((candidate) => {
			const existing = parseCellAddress(candidate.getAttribute('r') ?? '');
			return existing !== undefined && existing.col > address.col;
		});
		row.insertBefore(cell, after ?? null);
	}
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
function holds(cell: XmlElement, isNumeric: boolean, value: string): boolean {
	const type = cell.getAttribute('t') || 'n';
	if (isNumeric) {
		const current = childNamed(cell, 'v')?.textContent ?? undefined;
		if (value === '') return current === undefined && !childNamed(cell, 'is');
		if (type !== 'n' || current === undefined) return false;
		const before = Number.parseFloat(current);
		const after = Number.parseFloat(value);
		return Number.isFinite(before) && Number.isFinite(after) ? before === after : current === value;
	}
	return type === 'inlineStr' && inlineText(cell) === value;
}

/** Overwrite a cell's type and value to hold `value`, dropping any prior value shape. */
function writeCell(cell: XmlElement, isNumeric: boolean, value: string): void {
	for (const child of elements(cell))
		if (child.localName === 'is' || child.localName === 'v' || child.localName === 'f')
			cell.removeChild(child);
	if (isNumeric) {
		cell.removeAttribute('t');
		// A blank point empties the cell, as Excel stores a cleared cell.
		if (value === '') return;
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

/**
 * Apply a batch of cell writes to a parsed worksheet. Returns whether any cell changed, so the
 * caller only re-serializes a worksheet that differs from what was loaded.
 */
export function applyCellWritesToWorksheet(
	doc: XmlDocument,
	writes: readonly ChartCellWrite[],
): boolean {
	const sheetData = worksheetData(doc);
	if (!sheetData) return false;
	let changed = false;
	for (const write of writes) {
		if (!parseCellAddress(write.ref)) continue;
		const existing = findCell(sheetData, write.ref);
		if (existing && holds(existing, write.isNumeric, write.value)) continue;
		const cell = existing ?? ensureCell(sheetData, write.ref);
		if (!cell) continue;
		writeCell(cell, write.isNumeric, write.value);
		changed = true;
	}
	return changed;
}
