import { columnLabel } from '../../address.js';
import type { ErrorCode } from '../../model.js';
import { sheetIndex } from '../references.js';
import {
	ERR,
	fail,
	isError,
	LambdaValue,
	Matrix,
	RefValue,
	type Scalar,
	type Value,
} from '../values.js';
import { num, scalar, spec, str } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Information';

const ERROR_NUMBERS: Partial<Record<ErrorCode, number>> = {
	'#NULL!': 1,
	'#DIV/0!': 2,
	'#VALUE!': 3,
	'#REF!': 4,
	'#NAME?': 5,
	'#NUM!': 6,
	'#N/A': 7,
	'#GETTING_DATA': 8,
	'#SPILL!': 9,
	'#CALC!': 14,
};

const is = (
	name: string,
	syntax: string,
	description: string,
	test: (v: Scalar) => boolean,
): FunctionSpec => spec(name, C, syntax, description, 1, 1, (args) => test(scalar(args[0])));

const parity = (v: Scalar, odd: boolean): boolean => {
	if (isError(v)) fail(v);
	if (typeof v === 'boolean') fail(ERR.VALUE);
	const n = Math.trunc(num(v));
	return Math.abs(n % 2) === (odd ? 1 : 0);
};

function typeOf(value: Value | undefined): number {
	if (value instanceof RefValue) {
		if (!value.isCell()) return 64;
		return 0;
	}
	if (value instanceof Matrix) return 64;
	if (value instanceof LambdaValue) return 128;
	if (isError(value ?? null)) return 16;
	if (typeof value === 'number' || value === null || value === undefined) return 1;
	if (typeof value === 'string') return 2;
	return 4;
}

export const INFO_FUNCTIONS: FunctionSpec[] = [
	is('ISBLANK', 'ISBLANK(value)', 'TRUE when a cell is empty.', (v) => v === null),
	is('ISERROR', 'ISERROR(value)', 'TRUE for any error value.', (v) => isError(v)),
	is(
		'ISERR',
		'ISERR(value)',
		'TRUE for errors other than #N/A.',
		(v) => isError(v) && v.error !== '#N/A',
	),
	is('ISNA', 'ISNA(value)', 'TRUE for #N/A.', (v) => isError(v) && v.error === '#N/A'),
	is('ISNUMBER', 'ISNUMBER(value)', 'TRUE for numbers.', (v) => typeof v === 'number'),
	is('ISTEXT', 'ISTEXT(value)', 'TRUE for text.', (v) => typeof v === 'string'),
	is('ISNONTEXT', 'ISNONTEXT(value)', 'TRUE for anything but text.', (v) => typeof v !== 'string'),
	is('ISLOGICAL', 'ISLOGICAL(value)', 'TRUE for logical values.', (v) => typeof v === 'boolean'),
	is('ISEVEN', 'ISEVEN(number)', 'TRUE for even numbers.', (v) => parity(v, false)),
	is('ISODD', 'ISODD(number)', 'TRUE for odd numbers.', (v) => parity(v, true)),
	spec(
		'ISREF',
		C,
		'ISREF(value)',
		'TRUE for references.',
		1,
		1,
		(args) => args[0] instanceof RefValue,
		['any'],
	),
	spec(
		'ISFORMULA',
		C,
		'ISFORMULA(reference)',
		'TRUE when a cell holds a formula.',
		1,
		1,
		(args, ctx) => {
			const ref = args[0];
			if (!(ref instanceof RefValue)) return isError(ref ?? null) ? (ref as Scalar) : ERR.VALUE;
			const area = ref.areas[0];
			if (!area) return ERR.REF;
			return (
				ctx.frame.host.cellFormula(area.sheet, area.range.start.row, area.range.start.col) !==
				undefined
			);
		},
		['any'],
	),
	spec('N', C, 'N(value)', 'A value converted to a number (text is 0).', 1, 1, (args) => {
		const v = scalar(args[0]);
		if (isError(v)) return v;
		if (typeof v === 'number') return v;
		if (typeof v === 'boolean') return v ? 1 : 0;
		return 0;
	}),
	spec('NA', C, 'NA()', 'The #N/A error value.', 0, 0, () => ERR.NA),
	spec(
		'TYPE',
		C,
		'TYPE(value)',
		'The type of a value: 1 number, 2 text, 4 logical, 16 error, 64 array.',
		1,
		1,
		(args, ctx) => {
			const v = args[0];
			if (v instanceof RefValue && v.isCell()) return typeOf(ctx.toScalar(v));
			return typeOf(v);
		},
		['any'],
	),
	spec('ERROR.TYPE', C, 'ERROR.TYPE(error_val)', 'The number of an error value.', 1, 1, (args) => {
		const v = scalar(args[0]);
		return isError(v) ? (ERROR_NUMBERS[v.error] ?? ERR.NA) : ERR.NA;
	}),
	spec(
		'SHEET',
		C,
		'SHEET([value])',
		'The sheet number of a reference or sheet name.',
		0,
		1,
		(args, ctx) => {
			const v = args[0];
			if (v === undefined) return ctx.sheet + 1;
			if (v instanceof RefValue) return (v.areas[0]?.sheet ?? ctx.sheet) + 1;
			const index = sheetIndex(ctx.frame.host, str(v));
			return index < 0 ? ERR.NA : index + 1;
		},
		['any'],
	),
	spec(
		'SHEETS',
		C,
		'SHEETS([reference])',
		'The number of sheets in a reference (or the workbook).',
		0,
		1,
		(args, ctx) => {
			const v = args[0];
			if (v === undefined) return ctx.workbook.sheets.length;
			if (v instanceof RefValue) return new Set(v.areas.map((a) => a.sheet)).size;
			return ERR.VALUE;
		},
		['any'],
	),
	spec(
		'CELL',
		C,
		'CELL(info_type, [reference])',
		'Information about a cell (address, row, col, contents, type, filename, sheet).',
		1,
		2,
		(args, ctx) => {
			const info = str(ctx.toScalar(args[0] ?? null)).toLowerCase();
			const ref = args[1];
			let sheet = ctx.sheet;
			let row = ctx.row;
			let col = ctx.col;
			if (ref instanceof RefValue) {
				const area = ref.areas[0];
				if (!area) return ERR.REF;
				sheet = area.sheet;
				row = area.range.start.row;
				col = area.range.start.col;
			} else if (ref !== undefined) return ERR.VALUE;
			const value = ctx.readCell(sheet, row, col);
			switch (info) {
				case 'address':
					return `$${columnLabel(col)}$${row + 1}`;
				case 'row':
					return row + 1;
				case 'col':
					return col + 1;
				case 'contents':
					return value === null ? 0 : value;
				case 'type':
					return value === null ? 'b' : typeof value === 'string' ? 'l' : 'v';
				case 'filename':
				case 'sheet':
					return ctx.workbook.sheets[sheet]?.name ?? '';
				default:
					return ERR.VALUE;
			}
		},
		['any'],
		true,
	),
];
