import { formatAddress, formatRange } from '../address.js';
import { isCellError, type Cell, type RowInfo, type Worksheet } from '../model.js';
import { isSpilledCell } from '../formula/spill.js';
import { addFuturePrefixes } from '../read/formula-text.js';
import { SpillPlan } from './dynamic-array.js';
import type { SharedStringTable } from './shared-strings.js';
import { attrs, encodeEscapes, escapeText, num } from './xml-out.js';

/** Excel's longest cell text. */
const MAX_CELL_TEXT = 32_767;

export interface CellWriteContext {
	strings: SharedStringTable;
	styleCount: number;
	/** Text forced into table header cells (Excel requires headers to match column names). */
	headerText: ReadonlyMap<string, string>;
	/** Formulas forced into table totals-row cells. */
	formulas?: ReadonlyMap<string, string>;
	/** Set when a dynamic-array formula was written (the package then needs `xl/metadata.xml`). */
	dynamicArrays?: { used: boolean };
}

/** The `cm` value of the XLDAPR dynamic-array cell metadata block the writer emits. */
export const DYNAMIC_ARRAY_CM = 1;

const styleAttr = (ctx: CellWriteContext, styleId: number | undefined): number | undefined =>
	styleId && styleId > 0 && styleId < ctx.styleCount ? styleId : undefined;

function cellXml(
	ctx: CellWriteContext,
	spills: SpillPlan,
	source: Cell,
	row: number,
	col: number,
): string {
	const forced = ctx.formulas?.get(`${row}:${col}`);
	const cell: Cell = forced ? { ...source, formula: forced } : source;
	const r = formatAddress({ row, col });
	const s = styleAttr(ctx, cell.styleId);
	const header = ctx.headerText.get(`${row}:${col}`);
	if (header !== undefined)
		return `<c${attrs({ r, s, t: 's' })}><v>${ctx.strings.id(header)}</v></c>`;
	// A spilled value is written only inside the range its anchor records (as Excel does).
	if (isSpilledCell(source) && !forced && !spills.covered(source, row, col))
		return `<c${attrs({ r, s })}/>`;
	let f = '';
	let cm: number | undefined;
	if (cell.formula) {
		const formula = escapeText(addFuturePrefixes(cell.formula));
		const dynamic = forced ? undefined : spills.dynamicRange(cell, row, col);
		if (dynamic) {
			cm = DYNAMIC_ARRAY_CM;
			if (ctx.dynamicArrays) ctx.dynamicArrays.used = true;
			f = `<f t="array" ref="${formatRange(dynamic)}" aca="false">${formula}</f>`;
		} else
			f = cell.arrayRange
				? `<f t="array" ref="${formatRange(cell.arrayRange)}">${formula}</f>`
				: `<f>${formula}</f>`;
	}
	const value = cell.value;
	let t: string | undefined;
	let v = '';
	if (typeof value === 'number') {
		if (Number.isFinite(value)) v = num(value);
		else {
			t = 'e';
			v = '#NUM!';
		}
	} else if (typeof value === 'boolean') {
		t = 'b';
		v = value ? '1' : '0';
	} else if (typeof value === 'string') {
		const text = value.length > MAX_CELL_TEXT ? value.slice(0, MAX_CELL_TEXT) : value;
		if (cell.formula) {
			t = 'str';
			v = escapeText(encodeEscapes(text));
		} else {
			t = 's';
			v = String(ctx.strings.id(text, cell.richText));
		}
	} else if (isCellError(value)) {
		t = 'e';
		v = storedError(value.error);
	}
	if (!f && !v && t !== 'str') return `<c${attrs({ r, s })}/>`;
	return `<c${attrs({ r, s, t, cm })}>${f}${t === 'str' || v ? `<v>${v}</v>` : ''}</c>`;
}

/**
 * Error codes the file format stores. Excel writes `#SPILL!` and `#CALC!` as `#VALUE!` plus rich
 * value metadata and refuses a cell holding them as text; the cached value is only a placeholder
 * (the workbook is saved with full recalculation on load).
 */
const storedError = (code: string): string =>
	code === '#SPILL!' || code === '#CALC!' ? '#VALUE!' : code;

function rowAttrs(ctx: CellWriteContext, row: number, info: RowInfo | undefined): string {
	const s = styleAttr(ctx, info?.styleId);
	return attrs({
		r: row + 1,
		s,
		customFormat: s !== undefined ? true : undefined,
		ht: info?.height,
		hidden: info?.hidden || undefined,
		customHeight: info?.customHeight || undefined,
		outlineLevel: info?.outlineLevel || undefined,
		collapsed: info?.collapsed || undefined,
	});
}

/** `<sheetData>` in row and column order. */
export function sheetDataXml(ctx: CellWriteContext, sheet: Worksheet): string {
	const headerCols = new Map<number, number[]>();
	for (const key of [...ctx.headerText.keys(), ...(ctx.formulas?.keys() ?? [])]) {
		const [row = 0, col = 0] = key.split(':').map(Number);
		headerCols.set(row, [...(headerCols.get(row) ?? []), col]);
	}
	const rows = [
		...new Set([...sheet.rows.keys(), ...sheet.rowInfo.keys(), ...headerCols.keys()]),
	].sort((a, b) => a - b);
	const spills = new SpillPlan(sheet);
	let out = '';
	for (const row of rows) {
		const cells = sheet.rows.get(row);
		const info = sheet.rowInfo.get(row);
		const cols = [...new Set([...(cells?.keys() ?? []), ...(headerCols.get(row) ?? [])])].sort(
			(a, b) => a - b,
		);
		const body = cols
			.map((col) => cellXml(ctx, spills, cells?.get(col) ?? { value: null }, row, col))
			.join('');
		if (!body && !info) continue;
		out += body
			? `<row${rowAttrs(ctx, row, info)}>${body}</row>`
			: `<row${rowAttrs(ctx, row, info)}/>`;
	}
	return out ? `<sheetData>${out}</sheetData>` : '<sheetData/>';
}
