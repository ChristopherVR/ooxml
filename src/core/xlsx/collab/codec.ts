// JSON entries of the shared workbook and their conversion to and from the model. Everything a
// peer sends is validated on decode: a malformed entry decodes to undefined and is ignored.
import { formatRange, parseRange } from '../address';
import { isSpilledCell } from '../formula/spill';
import {
	type Cell,
	type ColumnInfo,
	type RichTextRun,
	type RowInfo,
	cellError,
	isCellError,
	isErrorCode,
} from '../model';
import { isRecord } from './schema';

/** A cell: value, type, formula, style key, array range, rich text and formula flags. */
export interface CellEntry {
	v?: number | string | boolean;
	/** `n` number, `s` string, `b` boolean, `e` error (`v` holds the code). */
	t?: 'n' | 's' | 'b' | 'e';
	f?: string;
	s?: string;
	a?: string;
	r?: RichTextRun[];
	d?: 1;
	l?: 1;
}

/** Row or column formatting: size, custom flag, hidden, style key, outline, collapsed. */
export interface LineEntry {
	size?: number;
	custom?: 1;
	hidden?: 1;
	s?: string;
	ol?: number;
	c?: 1;
	bf?: 1;
}

type KeyOf = (styleId: number | undefined) => string | undefined;
type IdOf = (key: string | undefined) => number | undefined;

const MAX_TEXT = 32_767;
const isText = (v: unknown, max = MAX_TEXT): v is string =>
	typeof v === 'string' && v.length <= max;
const isCount = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v) && v >= 0;

/** A cell's shared entry; undefined when it holds nothing worth sharing (spill results). */
export function encodeCell(cell: Cell, keyOf: KeyOf): CellEntry | undefined {
	const entry: CellEntry = {};
	const s = cell.styleId ? keyOf(cell.styleId) : undefined;
	if (isSpilledCell(cell)) return s ? { s } : undefined;
	const value = cell.value;
	if (isCellError(value)) {
		entry.v = value.error;
		entry.t = 'e';
	} else if (typeof value === 'number') {
		entry.v = value;
		entry.t = 'n';
	} else if (typeof value === 'string') {
		entry.v = value;
		entry.t = 's';
	} else if (typeof value === 'boolean') {
		entry.v = value;
		entry.t = 'b';
	}
	if (cell.formula !== undefined) entry.f = cell.formula;
	if (s) entry.s = s;
	if (cell.arrayRange) entry.a = formatRange(cell.arrayRange);
	if (cell.richText) entry.r = structuredClone(cell.richText);
	if (cell.dynamicArray) entry.d = 1;
	if (cell.legacyFormula) entry.l = 1;
	return Object.keys(entry).length ? entry : undefined;
}

function decodeRuns(raw: unknown): RichTextRun[] | undefined {
	if (!Array.isArray(raw) || raw.length > 1024) return undefined;
	const runs: RichTextRun[] = [];
	for (const run of raw) {
		if (!isRecord(run) || !isText(run.text)) return undefined;
		const out: RichTextRun = { text: run.text };
		if (isRecord(run.font)) out.font = structuredClone(run.font);
		runs.push(out);
	}
	return runs;
}

/** The model cell of a shared entry, or undefined when the entry is malformed. */
export function decodeCell(raw: unknown, idOf: IdOf): Cell | undefined {
	if (!isRecord(raw)) return undefined;
	const cell: Cell = { value: null };
	const { v, t } = raw;
	if (t === 'e') {
		if (typeof v !== 'string' || !isErrorCode(v)) return undefined;
		cell.value = cellError(v);
	} else if (t === 'n') {
		if (typeof v !== 'number' || !Number.isFinite(v)) return undefined;
		cell.value = v;
	} else if (t === 's') {
		if (!isText(v)) return undefined;
		cell.value = v;
	} else if (t === 'b') {
		if (typeof v !== 'boolean') return undefined;
		cell.value = v;
	} else if (t !== undefined || v !== undefined) return undefined;
	if (raw.f !== undefined) {
		if (!isText(raw.f, 8192)) return undefined;
		cell.formula = raw.f;
	}
	if (raw.s !== undefined && typeof raw.s !== 'string') return undefined;
	const styleId = idOf(raw.s as string | undefined);
	if (styleId) cell.styleId = styleId;
	if (typeof raw.a === 'string') {
		const range = parseRange(raw.a);
		if (range) cell.arrayRange = range;
	}
	if (raw.r !== undefined) {
		const runs = decodeRuns(raw.r);
		if (runs) cell.richText = runs;
	}
	if (raw.d === 1) cell.dynamicArray = true;
	if (raw.l === 1) cell.legacyFormula = true;
	return cell;
}

/** A row's or column's shared entry (an empty object keeps a bare span). */
export function encodeLine(info: RowInfo | ColumnInfo, keyOf: KeyOf, column = false): LineEntry {
	const entry: LineEntry = {};
	const size = column ? (info as ColumnInfo).width : (info as RowInfo).height;
	const custom = column ? (info as ColumnInfo).customWidth : (info as RowInfo).customHeight;
	if (size !== undefined) entry.size = size;
	if (custom) entry.custom = 1;
	if (info.hidden) entry.hidden = 1;
	const s = info.styleId ? keyOf(info.styleId) : undefined;
	if (s) entry.s = s;
	if (info.outlineLevel) entry.ol = info.outlineLevel;
	if (info.collapsed) entry.c = 1;
	if (column && (info as ColumnInfo).bestFit) entry.bf = 1;
	return entry;
}

function decodeLine(raw: unknown): LineEntry | undefined {
	if (!isRecord(raw)) return undefined;
	if (raw.size !== undefined && !isCount(raw.size)) return undefined;
	if (raw.ol !== undefined && !isCount(raw.ol)) return undefined;
	if (raw.s !== undefined && typeof raw.s !== 'string') return undefined;
	return raw as LineEntry;
}

export function decodeRow(raw: unknown, idOf: IdOf): RowInfo | undefined {
	const entry = decodeLine(raw);
	if (!entry) return undefined;
	const info: RowInfo = {};
	if (entry.size !== undefined) info.height = entry.size;
	if (entry.custom === 1) info.customHeight = true;
	if (entry.hidden === 1) info.hidden = true;
	const styleId = entry.s ? idOf(entry.s) : undefined;
	if (styleId) info.styleId = styleId;
	if (entry.ol) info.outlineLevel = entry.ol;
	if (entry.c === 1) info.collapsed = true;
	return info;
}

export function decodeColumn(
	raw: unknown,
	min: number,
	max: number,
	idOf: IdOf,
): ColumnInfo | undefined {
	const entry = decodeLine(raw);
	if (!entry) return undefined;
	const info: ColumnInfo = { min, max };
	if (entry.size !== undefined) info.width = entry.size;
	if (entry.custom === 1) info.customWidth = true;
	if (entry.hidden === 1) info.hidden = true;
	const styleId = entry.s ? idOf(entry.s) : undefined;
	if (styleId) info.styleId = styleId;
	if (entry.ol) info.outlineLevel = entry.ol;
	if (entry.c === 1) info.collapsed = true;
	if (entry.bf === 1) info.bestFit = true;
	return info;
}
