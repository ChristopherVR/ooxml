// Every edit in __fixtures__/word-table-structure.json was made in Word 16 and recorded as a grid;
// the engine must produce the same grid, or refuse the edits whose Word result is not a plain grid.
import { describe, expect, it } from 'vitest';
import fixture from './__fixtures__/word-table-structure.json';
import type { Table, TableCell } from './model';
import { tableGrid } from './table-grid';
import { deleteColumns, deleteRow, insertColumn, insertRow } from './table-structure';
import { mergeCells, splitCell, type GridRect } from './table-structure-merge';
import type { IdFactory } from './table-structure-cells';
import { twips } from './units';

interface Recorded {
	grid: number[];
	rows: { cells: string[]; trPr?: string[] }[];
}
interface Scenario {
	name: string;
	op: string;
	before: Recorded;
	after?: Recorded;
	wordError?: string;
}

const CELL = /^(.*?)(?::(\d+))?([\^|])?\{(\d+)\}$/;
let serial = 0;
const newId: IdFactory = () => `new-${++serial}`;

function cellOf(source: string, id: string): TableCell {
	const match = CELL.exec(source);
	if (!match) throw new Error(`Unreadable cell ${source}`);
	const [, text = '', span, merge, width = '0'] = match;
	const parts = text ? text.split('/') : [''];
	const cell: TableCell = {
		paragraphs: parts.map((part, i) => ({
			type: 'paragraph',
			id: `${id}p${i}`,
			runs: [{ text: part }],
		})),
		widthTwips: twips(Number(width)),
	};
	if (span) cell.gridSpan = Number(span);
	if (merge) cell.verticalMerge = merge === '^' ? 'restart' : 'continue';
	return cell;
}

function tableOf(recorded: Recorded): Table {
	const table: Table = {
		type: 'table',
		id: 't',
		grid: recorded.grid.map((w) => twips(w)),
		rows: recorded.rows.map((row, r) => row.cells.map((cell, c) => cellOf(cell, `r${r}c${c}`))),
	};
	const props = recorded.rows.map((row) => ({
		...(row.trPr?.includes('trHeight')
			? { heightTwips: twips(800), heightRule: 'atLeast' as const }
			: {}),
		...(row.trPr?.includes('tblHeader') ? { header: true } : {}),
	}));
	if (props.some((entry) => Object.keys(entry).length)) table.rowProperties = props;
	return table;
}

function describeCell(cell: TableCell): string {
	const text = cell.paragraphs
		.map((p) => p.runs.map((run) => run.text).join(''))
		.filter(Boolean)
		.join('/');
	return `${text}${cell.gridSpan && cell.gridSpan > 1 ? `:${cell.gridSpan}` : ''}${
		cell.verticalMerge === 'restart' ? '^' : cell.verticalMerge === 'continue' ? '|' : ''
	}{${cell.widthTwips}}`;
}

const rect = (firstRow: number, lastRow: number, firstCol: number, lastCol: number): GridRect => ({
	firstRow,
	lastRow,
	firstCol,
	lastCol,
});
type Edit = (t: Table) => Table | undefined;
/** The edit each recorded scenario performed, by name; `undefined` entries are refused on purpose. */
const edits: Record<string, Edit> = {
	'row-below-plain': (t) => insertRow(t, 1, 'below', newId),
	'row-above-plain': (t) => insertRow(t, 1, 'above', newId),
	'row-below-in-vmerge': (t) => insertRow(t, 0, 'below', newId),
	'row-above-in-vmerge': (t) => insertRow(t, 1, 'above', newId),
	'row-above-vmerge-top': (t) => insertRow(t, 0, 'above', newId),
	'row-below-vmerge-bottom': (t) => insertRow(t, 2, 'below', newId),
	'row-below-hmerge': (t) => insertRow(t, 0, 'below', newId),
	'row-above-hmerge': (t) => insertRow(t, 0, 'above', newId),
	'delete-row-vmerge-top': (t) => deleteRow(t, 0),
	'delete-row-vmerge-middle': (t) => deleteRow(t, 1),
	'delete-row-vmerge-pair': (t) => deleteRow(t, 2),
	'delete-row-vmerge-pair-top': (t) => deleteRow(t, 1),
	'delete-row-hmerge': (t) => deleteRow(t, 0),
	'col-right-plain': (t) => insertColumn(t, 1, 'right', newId),
	'col-left-plain': (t) => insertColumn(t, 1, 'left', newId),
	'col-right-through-hmerge': (t) => insertColumn(t, 0, 'right', newId),
	'col-right-from-hmerge': (t) => insertColumn(t, 0, 'right', newId),
	'col-right-vmerge': (t) => insertColumn(t, 0, 'right', newId),
	'col-delete-plain': (t) => deleteColumns(t, 1, 1),
	'col-delete-vmerge': (t) => deleteColumns(t, 1, 1),
	'merge-2x2': (t) => mergeCells(t, rect(0, 1, 0, 1), newId),
	'merge-row': (t) => mergeCells(t, rect(0, 0, 0, 2), newId),
	'merge-column': (t) => mergeCells(t, rect(0, 2, 0, 0), newId),
	'merge-with-empty': (t) => mergeCells(t, rect(0, 0, 0, 1), newId),
	'merge-into-vmerge': (t) => mergeCells(t, rect(0, 2, 1, 1), newId),
	'split-cols': (t) => splitCell(t, 1, 1, 1, 2, newId),
	'split-cols-3': (t) => splitCell(t, 1, 1, 1, 3, newId),
	'split-rows': (t) => splitCell(t, 1, 1, 2, 1, newId),
	'split-2x2': (t) => splitCell(t, 1, 1, 2, 2, newId),
	'split-hmerge': (t) => splitCell(t, 0, 0, 1, 2, newId),
	'split-vmerge': (t) => splitCell(t, 0, 0, 3, 1, newId),
};
/** Word's result for these is a table the model cannot hold (`gridAfter`, re-gridded rows) or an error. */
const refused: Record<string, Edit> = {
	'col-left-through-hmerge': (t) => insertColumn(t, 1, 'left', newId),
	'col-left-from-hmerge': (t) => insertColumn(t, 0, 'left', newId),
	'col-delete-through-hmerge': (t) => deleteColumns(t, 1, 1),
	'split-hmerge-3': (t) => splitCell(t, 0, 0, 1, 3, newId),
	'split-vmerge-2rows': (t) => splitCell(t, 0, 0, 2, 1, newId),
};

const scenarios = (fixture as { scenarios: Scenario[] }).scenarios;
const byName = new Map(scenarios.map((s) => [s.name, s]));

describe('table structure edits match Word', () => {
	it('covers every recorded scenario', () => {
		const handled = new Set([...Object.keys(edits), ...Object.keys(refused)]);
		const skipped = scenarios
			.map((s) => s.name)
			.filter((n) => !handled.has(n) && !n.startsWith('nested-'));
		expect(skipped).toEqual(['header-row-below']);
	});

	for (const [name, edit] of Object.entries(edits))
		it(`${name}: ${byName.get(name)?.op}`, () => {
			const scenario = byName.get(name)!;
			const result = edit(tableOf(scenario.before));
			expect(result, 'the edit is accepted').toBeDefined();
			expect(tableGrid(result!), 'the result tiles its grid').toBeDefined();
			expect(result!.grid?.map(Number)).toEqual(scenario.after!.grid);
			// New cells have no text, so only the cell text of the cells Word kept is compared.
			expect(result!.rows.map((row) => row.map(describeCell))).toEqual(
				scenario.after!.rows.map((row) => row.cells),
			);
		});

	for (const [name, edit] of Object.entries(refused))
		it(`${name}: refused (Word: ${byName.get(name)?.wordError ?? 'a result that is not a plain grid'})`, () => {
			expect(edit(tableOf(byName.get(name)!.before))).toBeUndefined();
		});

	it('header-row-below: the new row copies the row height and repeat-as-header flags', () => {
		const scenario = byName.get('header-row-below')!;
		const result = insertRow(tableOf(scenario.before), 0, 'below', newId)!;
		expect(
			result.rowProperties?.map((p) => [p.heightTwips !== undefined, p.header === true]),
		).toEqual([
			[true, true],
			[true, true],
			[false, false],
			[false, false],
		]);
	});
});
