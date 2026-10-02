// Create Names from Selection: one defined name per row or column of the selection, named from
// its label cell (top row, left column, bottom row or right column).
import {
	type CellRange,
	type DefinedName,
	formatAddress,
	getCell,
	quoteSheetName,
	validateDefinedName,
} from '@christophervr/xlsx-core';
import { guessHeader, target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { checkbox, fieldset } from './fields.js';
import { showDialog } from './frame.js';

export type LabelEdge = 'top' | 'left' | 'bottom' | 'right';

/** Turns label text into a valid name: spaces to `_`, other invalid characters dropped. */
export function sanitizeName(label: string): string | undefined {
	let name = label
		.trim()
		.replace(/\s+/g, '_')
		.replace(/[^A-Za-z0-9_.\\?]/g, '');
	if (!name) return undefined;
	if (!/^[A-Za-z_\\]/.test(name)) name = `_${name}`;
	if (validateDefinedName(name)) name = `_${name}`;
	return validateDefinedName(name) ? undefined : name;
}

const abs = (row: number, col: number): string =>
	formatAddress({ row, col }).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');

/** The names a selection yields for the chosen label edges. */
export function namesFromSelection(ctx: EditorContext, edges: readonly LabelEdge[]): DefinedName[] {
	const t = target(ctx);
	if (!t) return [];
	const r = t.range;
	const sheet = quoteSheetName(t.ws.name);
	const label = (row: number, col: number): string => {
		const value = getCell(t.ws, row, col)?.value;
		return value === null || value === undefined || typeof value === 'object' ? '' : String(value);
	};
	const ref = (range: CellRange): string => {
		const a = abs(range.start.row, range.start.col);
		const b = abs(range.end.row, range.end.col);
		return `${sheet}!${a === b ? a : `${a}:${b}`}`;
	};
	const top = edges.includes('top') ? 1 : 0;
	const bottom = edges.includes('bottom') ? 1 : 0;
	const left = edges.includes('left') ? 1 : 0;
	const right = edges.includes('right') ? 1 : 0;
	const out: DefinedName[] = [];
	const add = (text: string, range: CellRange): void => {
		const name = sanitizeName(text);
		if (!name || range.start.row > range.end.row || range.start.col > range.end.col) return;
		if (out.some((n) => n.name.toLowerCase() === name.toLowerCase())) return;
		out.push({ name, formula: ref(range) });
	};
	const rows = { start: r.start.row + top, end: r.end.row - bottom };
	const cols = { start: r.start.col + left, end: r.end.col - right };
	for (let col = cols.start; col <= cols.end; col++) {
		const range = { start: { row: rows.start, col }, end: { row: rows.end, col } };
		if (top) add(label(r.start.row, col), range);
		if (bottom) add(label(r.end.row, col), range);
	}
	for (let row = rows.start; row <= rows.end; row++) {
		const range = { start: { row, col: cols.start }, end: { row, col: cols.end } };
		if (left) add(label(row, r.start.col), range);
		if (right) add(label(row, r.end.col), range);
	}
	return out;
}

export function openCreateNames(ctx: EditorContext): Promise<number | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const header = guessHeader(t.ws, t.range);
	const edges: ReadonlyArray<readonly [LabelEdge, string, boolean]> = [
		['top', 'Top row', header],
		['left', 'Left column', !header],
		['bottom', 'Bottom row', false],
		['right', 'Right column', false],
	];
	const boxes = edges.map(([edge, label, on]) => ({ edge, ...checkbox(ctx, label, on) }));
	return showDialog<number>(ctx, {
		name: 'create-names',
		heading: 'Create Names from Selection',
		body: fieldset(ctx, 'Create names from values in the:', ...boxes.map((b) => b.wrapper)),
		opened: () => boxes[0]?.input.focus(),
		submit: () => {
			const chosen = boxes.filter((b) => b.input.checked).map((b) => b.edge);
			const names = namesFromSelection(ctx, chosen);
			if (!names.length) {
				ctx.toast(ctx.t('No names could be created from the selection.'), 'warning');
				return undefined;
			}
			t.session.batch('Create names', () => {
				for (const name of names) t.session.setDefinedName(name);
			});
			return names.length;
		},
	});
}
