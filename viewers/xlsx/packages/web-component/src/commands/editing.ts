// Home > Editing: AutoSum, Fill, Clear, Sort & Filter and Find & Select.
import { type CellRange, formatRange, getCell } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { icon } from './icons.js';
import { type Target, editing, target, viewing } from './util.js';
import { sortCommands } from './sort-filter.js';
import { selectSpecial } from './select-special.js';

export const AUTOSUM_FUNCTIONS: ReadonlyArray<readonly [fn: string, label: string]> = [
	['SUM', 'Sum'],
	['AVERAGE', 'Average'],
	['COUNT', 'Count Numbers'],
	['MAX', 'Max'],
	['MIN', 'Min'],
];

const isNumber = (t: Target, row: number, col: number): boolean =>
	typeof getCell(t.ws, row, col)?.value === 'number';

/** Excel's AutoSum guess for a single cell: the run of numbers above it, else to its left. */
export function autoSumRange(t: Target, row: number, col: number): CellRange | undefined {
	let r = row - 1;
	if (r >= 0 && isNumber(t, r, col)) {
		while (r - 1 >= 0 && isNumber(t, r - 1, col)) r--;
		return { start: { row: r, col }, end: { row: row - 1, col } };
	}
	let c = col - 1;
	if (c >= 0 && isNumber(t, row, c)) {
		while (c - 1 >= 0 && isNumber(t, row, c - 1)) c--;
		return { start: { row, col: c }, end: { row, col: col - 1 } };
	}
	return undefined;
}

function autoSum(ctx: EditorContext, fn: string): void {
	const t = target(ctx);
	if (!t) return;
	const r = t.range;
	const single = r.start.row === r.end.row && r.start.col === r.end.col;
	if (single) {
		const guess = autoSumRange(t, r.start.row, r.start.col);
		const formula = `=${fn}(${guess ? formatRange(guess) : ''})`;
		const grid = ctx.grid();
		if (grid) grid.beginEdit(formula);
		else t.session.setCellInput(t.sheet, r.start.row, r.start.col, formula);
		return;
	}
	// A block: one total under each column (in the row below, or the last row when it is empty).
	t.session.batch('AutoSum', () => {
		const lastEmpty = [...Array(r.end.col - r.start.col + 1).keys()].every(
			(i) => getCell(t.ws, r.end.row, r.start.col + i)?.value == null,
		);
		const row = lastEmpty ? r.end.row : r.end.row + 1;
		const bottom = lastEmpty ? r.end.row - 1 : r.end.row;
		for (let col = r.start.col; col <= r.end.col; col++) {
			const span: CellRange = { start: { row: r.start.row, col }, end: { row: bottom, col } };
			t.session.setCellInput(t.sheet, row, col, `=${fn}(${formatRange(span)})`);
		}
	});
}

/** The source band of a fill: the edge row / column, or the neighbour of a one-line selection. */
export function fillSource(
	r: CellRange,
	dir: 'down' | 'right' | 'up' | 'left',
): { source: CellRange; target: CellRange } | undefined {
	const { start, end } = r;
	const rows = end.row > start.row;
	const cols = end.col > start.col;
	switch (dir) {
		case 'down':
			if (rows) return { source: { start, end: { row: start.row, col: end.col } }, target: r };
			if (start.row === 0) return undefined;
			return {
				source: {
					start: { row: start.row - 1, col: start.col },
					end: { row: start.row - 1, col: end.col },
				},
				target: { start: { row: start.row - 1, col: start.col }, end },
			};
		case 'up':
			if (rows) return { source: { start: { row: end.row, col: start.col }, end }, target: r };
			return {
				source: {
					start: { row: end.row + 1, col: start.col },
					end: { row: end.row + 1, col: end.col },
				},
				target: { start, end: { row: end.row + 1, col: end.col } },
			};
		case 'right':
			if (cols) return { source: { start, end: { row: end.row, col: start.col } }, target: r };
			if (start.col === 0) return undefined;
			return {
				source: {
					start: { row: start.row, col: start.col - 1 },
					end: { row: end.row, col: start.col - 1 },
				},
				target: { start: { row: start.row, col: start.col - 1 }, end },
			};
		case 'left':
			if (cols) return { source: { start: { row: start.row, col: end.col }, end }, target: r };
			return {
				source: {
					start: { row: start.row, col: end.col + 1 },
					end: { row: end.row, col: end.col + 1 },
				},
				target: { start, end: { row: end.row, col: end.col + 1 } },
			};
	}
}

function fillDirection(ctx: EditorContext, dir: 'down' | 'right' | 'up' | 'left'): void {
	const t = target(ctx);
	if (!t) return;
	t.session.batch('Fill', () => {
		for (const r of t.ranges) {
			const spec = fillSource(r, dir);
			if (spec) t.session.fill(t.sheet, spec.source, spec.target);
		}
	});
}

export function editingCommands(): Command[] {
	return [
		...AUTOSUM_FUNCTIONS.map(([fn, label]) =>
			editing({
				id: fn === 'SUM' ? 'home.autosum' : `home.autosum-${fn.toLowerCase()}`,
				label: fn === 'SUM' ? 'AutoSum' : label,
				icon: icon('autoSum'),
				...(fn === 'SUM' ? { shortcut: 'Alt+=' } : {}),
				run: (ctx) => autoSum(ctx, fn),
			}),
		),
		editing({
			id: 'home.autosum-more',
			label: 'More Functions...',
			icon: icon('function'),
			run: (ctx) => void ctx.dialogs.open('insert-function'),
		}),
		editing({
			id: 'home.fill-down',
			label: 'Down',
			icon: icon('fillDown'),
			shortcut: 'Ctrl+D',
			run: (ctx) => fillDirection(ctx, 'down'),
		}),
		editing({
			id: 'home.fill-right',
			label: 'Right',
			icon: icon('fillRight'),
			shortcut: 'Ctrl+R',
			run: (ctx) => fillDirection(ctx, 'right'),
		}),
		editing({
			id: 'home.fill-up',
			label: 'Up',
			icon: icon('fillUp'),
			run: (ctx) => fillDirection(ctx, 'up'),
		}),
		editing({
			id: 'home.fill-left',
			label: 'Left',
			icon: icon('fillLeft'),
			run: (ctx) => fillDirection(ctx, 'left'),
		}),
		editing({
			id: 'home.fill-series',
			label: 'Series...',
			icon: icon('series'),
			run: (ctx) => void ctx.dialogs.open('fill-series'),
		}),
		...(
			[
				['all', 'Clear All'],
				['formats', 'Clear Formats'],
				['contents', 'Clear Contents'],
				['comments', 'Clear Comments and Notes'],
				['hyperlinks', 'Clear Hyperlinks'],
			] as const
		).map(([what, label]) =>
			editing({
				id: `home.clear-${what}`,
				label,
				icon: icon('clear'),
				...(what === 'formats' ? { lock: 'formatCells' } : {}),
				...(what === 'contents' ? { shortcut: 'Delete' } : {}),
				run: (ctx) => {
					const t = target(ctx);
					if (!t) return;
					t.session.batch(label, () => {
						for (const range of t.ranges) t.session.clearRange(t.sheet, range, what);
					});
				},
			}),
		),
		...sortCommands(),
		viewing({
			id: 'home.find',
			label: 'Find...',
			icon: icon('findSelect'),
			shortcut: 'Ctrl+F',
			run: (ctx) => void ctx.dialogs.open('find-replace', { tab: 'find' }),
		}),
		editing({
			id: 'home.replace',
			label: 'Replace...',
			icon: icon('findSelect'),
			shortcut: 'Ctrl+H',
			run: (ctx) => void ctx.dialogs.open('find-replace', { tab: 'replace' }),
		}),
		viewing({
			id: 'home.go-to',
			label: 'Go To...',
			icon: icon('goTo'),
			shortcut: 'Ctrl+G',
			run: (ctx) => void ctx.dialogs.open('go-to'),
		}),
		viewing({
			id: 'home.go-to-special',
			label: 'Go To Special...',
			icon: icon('goTo'),
			run: (ctx) => void ctx.dialogs.open('go-to-special'),
		}),
		...(
			[
				['formulas', 'Formulas'],
				['comments', 'Notes'],
				['conditional', 'Conditional Formatting'],
				['constants', 'Constants'],
				['validation', 'Data Validation'],
			] as const
		).map(([kind, label]) =>
			viewing({
				id: `home.select-${kind}`,
				label,
				icon: icon('findSelect'),
				run: (ctx) => {
					if (!selectSpecial(ctx, { kind })) ctx.toast(ctx.t('No cells were found.'), 'info');
				},
			}),
		),
	];
}
