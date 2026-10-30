import { type CellRange, parseAddress, parseRange } from '../address';
import type { SpilledCell } from '../formula/spill';
import { translateFormula } from '../formula/transform';
import {
	cellError,
	isErrorCode,
	type Cell,
	type DataTable,
	type RowInfo,
	type Worksheet,
} from '../model';
import { elements, type XmlElement } from '../../xml/index';
import { isoToSerial } from './dates';
import { stripFuturePrefixes } from './formula-text';
import { parseRichString, type SharedString } from './shared-strings';
import { att, boolAttr, numAttr, xFirst, xText } from './xml-util';

export interface CellContext {
	sharedStrings: readonly SharedString[];
	xfMap: readonly number[];
	date1904: boolean;
	palette: readonly string[] | undefined;
	/** One-based `cm` values that mark dynamic-array formulas (from `xl/metadata.xml`). */
	dynamicCells?: ReadonlySet<number>;
	warn(message: string): void;
}

/** A dynamic-array anchor read from the file, with the spill range Excel recorded. */
interface DynamicAnchor {
	row: number;
	col: number;
	range: CellRange;
}

interface SharedMaster {
	row: number;
	col: number;
	formula: string;
}

const styleOf = (ctx: CellContext, element: XmlElement): number => {
	const s = numAttr(element, 's');
	return s === undefined ? 0 : (ctx.xfMap[s] ?? 0);
};

function readValue(ctx: CellContext, c: XmlElement, cell: Cell): void {
	const type = att(c, 't') ?? 'n';
	const v = xFirst(c, 'v');
	const raw = v ? (v.textContent ?? '') : undefined;
	switch (type) {
		case 's': {
			const entry = ctx.sharedStrings[Number(raw)];
			if (!entry) {
				if (raw !== undefined) ctx.warn('A cell refers to a missing shared string.');
				return;
			}
			cell.value = entry.text;
			if (entry.runs) cell.richText = entry.runs.map((run) => ({ ...run }));
			if (entry.phonetic) cell.phonetic = structuredClone(entry.phonetic);
			return;
		}
		case 'inlineStr': {
			const is = xFirst(c, 'is');
			if (!is) return;
			const entry = parseRichString(is, ctx.palette);
			cell.value = entry.text;
			if (entry.runs) cell.richText = entry.runs;
			if (entry.phonetic) cell.phonetic = entry.phonetic;
			return;
		}
		case 'str':
			if (v) cell.value = xText(v);
			return;
		case 'b':
			if (raw !== undefined) cell.value = raw.trim() === '1' || raw.trim().toLowerCase() === 'true';
			return;
		case 'e': {
			const code = (raw ?? '').trim();
			if (raw !== undefined) cell.value = cellError(isErrorCode(code) ? code : '#VALUE!');
			return;
		}
		case 'd': {
			if (raw === undefined) return;
			const serial = isoToSerial(raw, ctx.date1904);
			if (serial === undefined) ctx.warn(`Unreadable date cell value "${raw}".`);
			else cell.value = serial;
			return;
		}
		default: {
			if (raw === undefined || raw.trim() === '') return;
			const n = Number(raw);
			cell.value = Number.isFinite(n) ? n : raw;
		}
	}
}

function readDataTable(f: XmlElement): DataTable | undefined {
	const ref = parseRange(att(f, 'ref') ?? '');
	const r1 = att(f, 'r1');
	if (!ref || !r1) return undefined;
	const table: DataTable = { ref, r1 };
	const r2 = att(f, 'r2');
	if (r2) table.r2 = r2;
	if (boolAttr(f, 'dt2D')) table.dt2D = true;
	if (boolAttr(f, 'dtr')) table.dtr = true;
	if (boolAttr(f, 'del1')) table.del1 = true;
	if (boolAttr(f, 'del2')) table.del2 = true;
	return table;
}

function readFormula(
	ctx: CellContext,
	f: XmlElement,
	row: number,
	col: number,
	cell: Cell,
	masters: Map<string, SharedMaster>,
	dynamic: boolean,
	anchors: DynamicAnchor[],
	sheet: Worksheet,
): void {
	const kind = att(f, 't') ?? 'normal';
	if (dynamic) cell.dynamicArray = true;
	else if (kind !== 'array') cell.legacyFormula = true;
	const text = stripFuturePrefixes(f.textContent ?? '');
	if (kind === 'shared') {
		const si = att(f, 'si') ?? '';
		if (text.trim()) {
			masters.set(si, { row, col, formula: text });
			cell.formula = text;
			return;
		}
		const master = masters.get(si);
		if (!master) {
			ctx.warn('A shared formula refers to a missing master cell; its formula was dropped.');
			return;
		}
		cell.formula = translateFormula(master.formula, row - master.row, col - master.col);
		return;
	}
	if (kind === 'dataTable') {
		delete cell.legacyFormula;
		const table = readDataTable(f);
		if (table) (sheet.dataTables ??= []).push(table);
		else ctx.warn('A what-if data table formula could not be read; its cached values were kept.');
		return;
	}
	if (!text.trim()) return;
	cell.formula = text;
	if (kind === 'array') {
		const range = parseRange(att(f, 'ref') ?? '');
		if (dynamic) {
			if (range) anchors.push({ row, col, range });
		} else if (range) cell.arrayRange = range;
	}
}

function readRowInfo(ctx: CellContext, rowEl: XmlElement): RowInfo | undefined {
	const info: RowInfo = {};
	const height = numAttr(rowEl, 'ht');
	if (height !== undefined) info.height = height;
	if (boolAttr(rowEl, 'customHeight')) info.customHeight = true;
	if (boolAttr(rowEl, 'hidden')) info.hidden = true;
	if (boolAttr(rowEl, 'customFormat')) {
		const styleId = styleOf(ctx, rowEl);
		if (styleId) info.styleId = styleId;
	}
	const outline = numAttr(rowEl, 'outlineLevel');
	if (outline) info.outlineLevel = outline;
	if (boolAttr(rowEl, 'collapsed')) info.collapsed = true;
	if (boolAttr(rowEl, 'thickTop')) info.thickTop = true;
	if (boolAttr(rowEl, 'thickBot')) info.thickBot = true;
	return Object.keys(info).length ? info : undefined;
}

/**
 * Reads the children of `<sheetData>` (`elements(sheetData)`, or the rows of a split part) into
 * the sheet's cell map and row info.
 */
export function readSheetData(
	ctx: CellContext,
	rows: Iterable<XmlElement>,
	sheet: Worksheet,
): void {
	const reader = sheetDataReader(ctx, sheet);
	for (const rowEl of rows) reader.read(rowEl);
	reader.done();
}

/**
 * Reads `<sheetData>` children into `sheet` one at a time, carrying shared formulas and the row
 * position from one to the next; `done` marks the cells dynamic arrays spill into.
 */
export function sheetDataReader(ctx: CellContext, sheet: Worksheet) {
	const masters = new Map<string, SharedMaster>();
	const anchors: DynamicAnchor[] = [];
	let rowIndex = -1;
	const read = (rowEl: XmlElement): void => {
		if (rowEl.localName !== 'row') return;
		const r = numAttr(rowEl, 'r');
		rowIndex = r !== undefined ? r - 1 : rowIndex + 1;
		const info = readRowInfo(ctx, rowEl);
		if (info) sheet.rowInfo.set(rowIndex, info);
		let colIndex = -1;
		let cells: Map<number, Cell> | undefined;
		for (const c of elements(rowEl)) {
			if (c.localName !== 'c') continue;
			const ref = att(c, 'r');
			const address = ref ? parseAddress(ref) : undefined;
			colIndex = address ? address.col : colIndex + 1;
			const row = address ? address.row : rowIndex;
			const cell: Cell = { value: null };
			const styleId = styleOf(ctx, c);
			if (styleId) cell.styleId = styleId;
			readValue(ctx, c, cell);
			const f = xFirst(c, 'f');
			const cm = numAttr(c, 'cm');
			const dynamic = cm !== undefined && (ctx.dynamicCells?.has(cm) ?? false);
			if (cm !== undefined && !dynamic) cell.cellMetadata = cm;
			const vm = numAttr(c, 'vm');
			if (vm !== undefined) cell.valueMetadata = vm;
			if (f) readFormula(ctx, f, row, colIndex, cell, masters, dynamic, anchors, sheet);
			if (cell.formula === undefined) {
				delete cell.legacyFormula;
				delete cell.dynamicArray;
			}
			if (cell.value === null && !cell.formula && !cell.styleId) continue;
			if (row !== rowIndex) {
				let other = sheet.rows.get(row);
				if (!other) sheet.rows.set(row, (other = new Map()));
				other.set(colIndex, cell);
				continue;
			}
			if (!cells) {
				cells = sheet.rows.get(rowIndex) ?? new Map();
				sheet.rows.set(rowIndex, cells);
			}
			cells.set(colIndex, cell);
		}
	};
	return { read, done: () => markSpills(sheet, anchors) };
}

/**
 * Cells Excel saved inside a dynamic array's spill range hold cached results, not user data:
 * they are marked as spilled so the calc engine can spill over them again.
 */
function markSpills(sheet: Worksheet, anchors: readonly DynamicAnchor[]): void {
	for (const anchor of anchors) {
		const { start, end } = anchor.range;
		// Visit only populated rows and cells: a ref such as A1:XFD1048576 must not cost a million rows.
		const mark = (r: number, cells: Map<number, Cell>): void => {
			const visit = (c: number, cell: Cell | undefined): void => {
				if (c < start.col || c > end.col || (r === anchor.row && c === anchor.col)) return;
				if (!cell || cell.formula !== undefined) return;
				(cell as SpilledCell).spillAnchor = { row: anchor.row, col: anchor.col };
			};
			if (end.col - start.col + 1 <= cells.size) {
				for (let c = start.col; c <= end.col; c++) visit(c, cells.get(c));
			} else for (const [c, cell] of cells) visit(c, cell);
		};
		if (end.row - start.row + 1 <= sheet.rows.size) {
			for (let r = start.row; r <= end.row; r++) {
				const cells = sheet.rows.get(r);
				if (cells) mark(r, cells);
			}
		} else {
			for (const [r, cells] of sheet.rows) if (r >= start.row && r <= end.row) mark(r, cells);
		}
	}
}
