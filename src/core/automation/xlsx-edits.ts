import { columnIndex, MAX_COL, MAX_ROW, parseAddress, parseRange } from '../xlsx/address';
import { createEditSession } from '../xlsx/edit/session';
import type { ClearWhat, MergeMode } from '../xlsx/edit/types';
import type { Workbook } from '../xlsx/model';
import type { StylePatch } from '../xlsx/styles';
import { validateSheetName } from '../xlsx/workbook';

/**
 * One workbook edit, as plain JSON an agent can send and a reviewer can read. Sheets are zero-based
 * indexes, cells and ranges are A1 references, rows are numbered from 1 and columns are letters,
 * as Excel shows them. Edits run in order, each on the workbook the previous ones left, so a
 * sheet index or row number refers to the workbook at that point.
 */
export type XlsxEdit =
	/** Cell inputs as typed into the grid: `=` starts a formula, `'` forces text. */
	| { op: 'setCells'; sheet: number; cells: { address: string; input: string }[] }
	/** Clears contents (the default), formats, comments, hyperlinks or all of them. */
	| { op: 'clear'; sheet: number; range: string; what?: ClearWhat }
	| { op: 'format'; sheet: number; range: string; style: StylePatch }
	| { op: 'merge'; sheet: number; range: string; mode?: MergeMode }
	| { op: 'unmerge'; sheet: number; range: string }
	/** Inserts `count` rows above `row`, or deletes `count` rows from `row` down. */
	| { op: 'insertRows' | 'deleteRows'; sheet: number; row: number; count: number }
	/** Inserts `count` columns left of `column`, or deletes `count` columns from it rightwards. */
	| { op: 'insertColumns' | 'deleteColumns'; sheet: number; column: string; count: number }
	| { op: 'addSheet'; name?: string; at?: number }
	| { op: 'renameSheet'; sheet: number; name: string }
	| { op: 'deleteSheet'; sheet: number }
	/** Adds or replaces a defined name; `sheet` scopes it to that sheet. A leading `=` is optional. */
	| { op: 'setName'; name: string; formula: string; sheet?: number }
	| { op: 'deleteName'; name: string; sheet?: number };

export const XLSX_EDIT_LIMITS = { edits: 1000, cells: 10_000 } as const;

const CLEAR: readonly ClearWhat[] = ['all', 'contents', 'formats', 'comments', 'hyperlinks'];
const MERGE: readonly MergeMode[] = ['merge', 'center', 'across'];
const STYLE_KEYS = new Set(['font', 'fill', 'border', 'numFmt', 'alignment', 'protection']);

function fail(index: number, message: string): never {
	throw new Error(`Edit ${index + 1}: ${message}`);
}

/**
 * Applies `edits` to `workbook` as one undo step through the core edit session (formulas are
 * recalculated). Throws, naming the edit, on the first one that is invalid; the workbook is then
 * rolled back. Rows and columns are matched by address in a comparison, so an inserted row
 * shows as a change to every cell below it.
 */
export function applyXlsxEdits(workbook: Workbook, edits: readonly XlsxEdit[]): void {
	if (!Array.isArray(edits)) throw new Error('Edits must be a list');
	if (edits.length > XLSX_EDIT_LIMITS.edits)
		throw new Error(`Make at most ${XLSX_EDIT_LIMITS.edits} edits per request`);
	const cellCount = edits.reduce(
		(sum, edit) =>
			sum + (edit?.op === 'setCells' && Array.isArray(edit.cells) ? edit.cells.length : 0),
		0,
	);
	if (cellCount > XLSX_EDIT_LIMITS.cells)
		throw new Error(`Edit at most ${XLSX_EDIT_LIMITS.cells} cells per request`);
	const session = createEditSession(workbook);
	session.batch('Automation edits', () => {
		edits.forEach((edit, index) => applyEdit(workbook, session, edit, index));
	});
}

type Session = ReturnType<typeof createEditSession>;

function applyEdit(workbook: Workbook, session: Session, edit: XlsxEdit, index: number) {
	const sheet = (value: unknown) => {
		if (!Number.isSafeInteger(value) || (value as number) < 0 || !workbook.sheets[value as number])
			fail(index, 'Sheet does not exist');
		return value as number;
	};
	const range = (value: unknown) => {
		const parsed = typeof value === 'string' ? parseRange(value) : undefined;
		if (!parsed || parsed.end.row > MAX_ROW || parsed.end.col > MAX_COL)
			fail(index, `Invalid range: ${String(value)}`);
		return parsed;
	};
	const count = (value: unknown, at: number, max: number) => {
		if (!Number.isSafeInteger(value) || (value as number) < 1 || at + (value as number) > max + 1)
			fail(index, 'Invalid count');
		return value as number;
	};
	const name = (value: unknown) => {
		if (typeof value !== 'string' || !value) fail(index, 'A name is required');
		return value;
	};
	switch (edit?.op) {
		case 'setCells': {
			const s = sheet(edit.sheet);
			if (!Array.isArray(edit.cells)) fail(index, 'Cells must be a list');
			for (const cell of edit.cells) {
				const at = parseAddress(String(cell?.address));
				if (!at) fail(index, `Invalid cell address: ${String(cell?.address)}`);
				if (typeof cell.input !== 'string') fail(index, 'A cell input must be text');
				session.setCellInput(s, at.row, at.col, cell.input);
			}
			return;
		}
		case 'clear': {
			const what = edit.what ?? 'contents';
			if (!CLEAR.includes(what)) fail(index, `Cannot clear ${String(what)}`);
			session.clearRange(sheet(edit.sheet), range(edit.range), what);
			return;
		}
		case 'format': {
			const style = edit.style as unknown;
			if (!style || typeof style !== 'object' || Array.isArray(style))
				fail(index, 'A style patch is required');
			const unknown = Object.keys(style).filter((key) => !STYLE_KEYS.has(key));
			if (unknown.length) fail(index, `Unknown style keys: ${unknown.join(', ')}`);
			session.applyStyle(sheet(edit.sheet), [range(edit.range)], style as StylePatch);
			return;
		}
		case 'merge': {
			const mode = edit.mode ?? 'merge';
			if (!MERGE.includes(mode)) fail(index, `Unknown merge mode: ${String(mode)}`);
			session.merge(sheet(edit.sheet), range(edit.range), mode);
			return;
		}
		case 'unmerge':
			session.unmerge(sheet(edit.sheet), range(edit.range));
			return;
		case 'insertRows':
		case 'deleteRows': {
			const s = sheet(edit.sheet);
			if (!Number.isSafeInteger(edit.row) || edit.row < 1 || edit.row > MAX_ROW + 1)
				fail(index, 'Invalid row');
			const at = edit.row - 1;
			const n = count(edit.count, at, MAX_ROW);
			if (edit.op === 'insertRows') session.insertRows(s, at, n);
			else session.deleteRows(s, at, n);
			return;
		}
		case 'insertColumns':
		case 'deleteColumns': {
			const s = sheet(edit.sheet);
			const at =
				typeof edit.column === 'string' && /^[A-Za-z]{1,3}$/.test(edit.column)
					? columnIndex(edit.column)
					: -1;
			if (at < 0 || at > MAX_COL) fail(index, `Invalid column: ${String(edit.column)}`);
			const n = count(edit.count, at, MAX_COL);
			if (edit.op === 'insertColumns') session.insertColumns(s, at, n);
			else session.deleteColumns(s, at, n);
			return;
		}
		case 'addSheet': {
			if (edit.name !== undefined) {
				const problem = validateSheetName(workbook, name(edit.name));
				if (problem) fail(index, problem);
			}
			const at = edit.at;
			if (at !== undefined && (!Number.isSafeInteger(at) || at < 0 || at > workbook.sheets.length))
				fail(index, 'Invalid sheet position');
			session.addSheet(edit.name, at);
			return;
		}
		case 'renameSheet': {
			const s = sheet(edit.sheet);
			const problem = validateSheetName(workbook, name(edit.name), s);
			if (problem) fail(index, problem);
			session.renameSheet(s, edit.name);
			return;
		}
		case 'deleteSheet': {
			const s = sheet(edit.sheet);
			if (workbook.sheets.length < 2) fail(index, 'A workbook must keep at least one sheet');
			session.deleteSheet(s);
			return;
		}
		case 'setName': {
			if (typeof edit.formula !== 'string' || !edit.formula) fail(index, 'A formula is required');
			session.setDefinedName({
				name: name(edit.name),
				formula: edit.formula,
				...(edit.sheet !== undefined ? { localSheet: sheet(edit.sheet) } : {}),
			});
			return;
		}
		case 'deleteName':
			session.deleteDefinedName(
				name(edit.name),
				edit.sheet !== undefined ? sheet(edit.sheet) : undefined,
			);
			return;
		default:
			fail(index, `Unknown edit: ${String((edit as { op?: unknown } | undefined)?.op)}`);
	}
}
