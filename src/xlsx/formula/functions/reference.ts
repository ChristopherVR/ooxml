import { columnLabel, MAX_COL, MAX_ROW, quoteSheetName } from '../../address.js';
import type { CallContext } from '../context.js';
import { sheetIndex } from '../references.js';
import { ERR, fail, isError, Matrix, RefValue, type Scalar, type Value } from '../values.js';
import { bool, int, optNum, spec, str } from './helpers.js';
import { shape } from './lookup-core.js';
import type { FunctionSpec } from './types.js';

const C = 'Lookup & Reference';

function refArg(value: Value | undefined): RefValue {
	if (value instanceof RefValue) return value;
	if (isError(value as Scalar)) fail(value as never);
	return fail(ERR.VALUE);
}

function offset(args: Value[]): Value {
	const ref = refArg(args[0]);
	const area = ref.areas[0];
	if (!area || ref.areas.length !== 1) fail(ERR.VALUE);
	const { start, end } = area.range;
	const height = args.length > 3 && args[3] !== null ? int(args[3]) : end.row - start.row + 1;
	const width = args.length > 4 && args[4] !== null ? int(args[4]) : end.col - start.col + 1;
	if (height === 0 || width === 0) fail(ERR.REF);
	const top = start.row + int(args[1]);
	const left = start.col + int(args[2]);
	const r1 = height > 0 ? top : top + height + 1;
	const r2 = height > 0 ? top + height - 1 : top;
	const c1 = width > 0 ? left : left + width + 1;
	const c2 = width > 0 ? left + width - 1 : left;
	if (r1 < 0 || c1 < 0 || r2 > MAX_ROW || c2 > MAX_COL) fail(ERR.REF);
	return new RefValue([
		{ sheet: area.sheet, range: { start: { row: r1, col: c1 }, end: { row: r2, col: c2 } } },
	]);
}

/** Parses an R1C1 reference relative to the formula cell. */
function parseR1C1(text: string, ctx: CallContext): RefValue | undefined {
	const bang = text.lastIndexOf('!');
	let sheet = ctx.sheet;
	let body = text;
	if (bang >= 0) {
		const name = text
			.slice(0, bang)
			.replace(/^'(.*)'$/, '$1')
			.replace(/''/g, "'");
		sheet = sheetIndex(ctx.frame.host, name);
		if (sheet < 0) return undefined;
		body = text.slice(bang + 1);
	}
	const part = (spec: string | undefined, base: number): number | undefined => {
		if (spec === undefined || spec === '') return base;
		if (spec.startsWith('[')) return base + Number(spec.slice(1, -1));
		return Number(spec) - 1;
	};
	const cells = body.split(':').map((p) => /^R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?$/i.exec(p));
	if (cells.length > 2 || cells.some((m) => !m)) return undefined;
	const corners = cells.map((m) => ({ row: part(m?.[1], ctx.row), col: part(m?.[2], ctx.col) }));
	const a = corners[0];
	const b = corners[corners.length - 1];
	if (
		!a ||
		!b ||
		a.row === undefined ||
		a.col === undefined ||
		b.row === undefined ||
		b.col === undefined
	)
		return undefined;
	const range = {
		start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
		end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
	};
	if (
		range.start.row < 0 ||
		range.start.col < 0 ||
		range.end.row > MAX_ROW ||
		range.end.col > MAX_COL
	)
		return undefined;
	return new RefValue([{ sheet, range }]);
}

function address(args: Value[]): string {
	const row = int(args[0]);
	const col = int(args[1]);
	const abs = Math.trunc(optNum(args, 2, 1));
	const a1 = args.length > 3 && args[3] !== null ? bool(args[3]) : true;
	if (row < 1 || col < 1 || row > MAX_ROW + 1 || col > MAX_COL + 1 || abs < 1 || abs > 4)
		fail(ERR.VALUE);
	const absRow = abs === 1 || abs === 2;
	const absCol = abs === 1 || abs === 3;
	let ref: string;
	if (a1) ref = `${absCol ? '$' : ''}${columnLabel(col - 1)}${absRow ? '$' : ''}${row}`;
	else ref = `R${absRow ? row : `[${row}]`}C${absCol ? col : `[${col}]`}`;
	if (args.length > 4 && args[4] !== null) {
		const sheet = str(args[4]);
		if (sheet !== '') ref = `${quoteSheetName(sheet)}!${ref}`;
	}
	return ref;
}

function rowsOrCols(value: Value | undefined, ctx: CallContext, axis: 'row' | 'col'): Value {
	if (value === undefined) return (axis === 'row' ? ctx.row : ctx.col) + 1;
	const ref = refArg(value);
	const area = ref.areas[0];
	if (!area) fail(ERR.REF);
	const { start, end } = area.range;
	if (axis === 'row') {
		if (start.row === end.row) return start.row + 1;
		return Matrix.build(end.row - start.row + 1, 1, (r) => start.row + r + 1);
	}
	if (start.col === end.col) return start.col + 1;
	return Matrix.build(1, end.col - start.col + 1, (_r, c) => start.col + c + 1);
}

export const REFERENCE_FUNCTIONS: FunctionSpec[] = [
	spec(
		'OFFSET',
		C,
		'OFFSET(reference, rows, cols, [height], [width])',
		'A reference offset from a starting reference.',
		3,
		5,
		(args) => offset(args),
		['any', 'value'],
		true,
	),
	spec(
		'INDIRECT',
		C,
		'INDIRECT(ref_text, [a1])',
		'The reference named by a text.',
		1,
		2,
		(args, ctx) => {
			const text = str(args[0]).trim();
			const a1 = args.length > 1 && args[1] !== null ? bool(args[1]) : true;
			const ref = a1 ? ctx.parseReference(text) : parseR1C1(text, ctx);
			return ref ?? ERR.REF;
		},
		undefined,
		true,
	),
	spec(
		'ROW',
		C,
		'ROW([reference])',
		'The row number of a reference.',
		0,
		1,
		(args, ctx) => rowsOrCols(args[0], ctx, 'row'),
		['any'],
	),
	spec(
		'COLUMN',
		C,
		'COLUMN([reference])',
		'The column number of a reference.',
		0,
		1,
		(args, ctx) => rowsOrCols(args[0], ctx, 'col'),
		['any'],
	),
	spec(
		'ROWS',
		C,
		'ROWS(array)',
		'The number of rows in a reference or array.',
		1,
		1,
		(args) => {
			if (isError(args[0] as Scalar)) return args[0] ?? null;
			return shape(args[0] ?? null).rows;
		},
		['any'],
	),
	spec(
		'COLUMNS',
		C,
		'COLUMNS(array)',
		'The number of columns in a reference or array.',
		1,
		1,
		(args) => {
			if (isError(args[0] as Scalar)) return args[0] ?? null;
			return shape(args[0] ?? null).cols;
		},
		['any'],
	),
	spec(
		'AREAS',
		C,
		'AREAS(reference)',
		'The number of areas in a reference.',
		1,
		1,
		(args) => refArg(args[0]).areas.length,
		['any'],
	),
	spec(
		'ADDRESS',
		C,
		'ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])',
		'A cell address as text.',
		2,
		5,
		(args) => address(args),
	),
	spec(
		'HYPERLINK',
		C,
		'HYPERLINK(link_location, [friendly_name])',
		'A link; the cell shows the friendly name.',
		1,
		2,
		(args) => (args.length > 1 && args[1] !== null ? (args[1] as Scalar) : str(args[0])),
	),
	spec(
		'ANCHORARRAY',
		C,
		'ANCHORARRAY(reference)',
		'The spill range of a dynamic array formula (A1#).',
		1,
		1,
		(args, ctx) => {
			const ref = refArg(args[0]);
			const area = ref.areas[0];
			if (!area) return ERR.REF;
			const spill = ctx.frame.host.spillRange(
				area.sheet,
				area.range.start.row,
				area.range.start.col,
			);
			return spill ? new RefValue([{ sheet: area.sheet, range: spill }]) : ERR.REF;
		},
		['any'],
	),
	spec(
		'SINGLE',
		C,
		'SINGLE(value)',
		'Implicit intersection (the @ operator).',
		1,
		1,
		(args, ctx) => ctx.toScalar(args[0] ?? null),
		['any'],
	),
	spec(
		'FORMULATEXT',
		C,
		'FORMULATEXT(reference)',
		'The formula of a cell as text.',
		1,
		1,
		(args, ctx) => {
			const area = refArg(args[0]).areas[0];
			if (!area) return ERR.REF;
			const formula = ctx.frame.host.cellFormula(
				area.sheet,
				area.range.start.row,
				area.range.start.col,
			);
			return formula === undefined ? ERR.NA : `=${formula}`;
		},
		['any'],
	),
];
