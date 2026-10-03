/** The selection model behind `ctx.selection`, and A1 text conversions for the element API. */
import {
	formatAddress,
	formatRange,
	MAX_COL,
	MAX_ROW,
	normalizeRange,
	parseRange,
	type CellAddress,
	type CellRange,
	type Workbook,
	type Worksheet,
} from '@christophervr/xlsx-core';
import type { Selection, SelectionModel } from './context';

const A1: CellAddress = { row: 0, col: 0 };
const cellRange = (at: CellAddress): CellRange => ({ start: { ...at }, end: { ...at } });

export function initialSelection(sheet: number, workbook?: Workbook): Selection {
	const saved = workbook?.sheets[sheet]?.view?.selection;
	const active = saved?.active ?? A1;
	const ranges = saved?.ranges?.length ? saved.ranges.map(normalizeRange) : [cellRange(active)];
	return { sheet, active: { ...active }, anchor: { ...active }, ranges };
}

const same = (a: Selection, b: Selection) => JSON.stringify(a) === JSON.stringify(b);

export function createSelectionModel(initial: Selection = initialSelection(0)): SelectionModel {
	let current = initial;
	const listeners = new Set<(selection: Selection) => void>();
	return {
		get: () => current,
		set(next) {
			const active =
				next.active ?? (next.ranges?.[0]?.start ? { ...next.ranges[0].start } : current.active);
			const merged: Selection = {
				sheet: next.sheet ?? current.sheet,
				active,
				anchor: next.anchor ?? (next.active || next.ranges ? { ...active } : current.anchor),
				ranges: (next.ranges ?? (next.active ? [cellRange(active)] : current.ranges)).map(
					normalizeRange,
				),
			};
			const cellsMoved =
				next.active !== undefined || next.ranges !== undefined || next.sheet !== undefined;
			const drawing = 'drawing' in next ? next.drawing : cellsMoved ? undefined : current.drawing;
			if (drawing !== undefined) merged.drawing = drawing;
			if (same(merged, current)) return;
			current = merged;
			for (const listener of [...listeners]) listener(current);
		},
		onChange(listener) {
			listeners.add(listener);
			return () => {
				listeners.delete(listener);
			};
		},
	};
}

/**
 * Per-sheet selection memory, keyed by the worksheet object so it survives sheets being
 * inserted, moved or deleted around it. Switching tabs comes back to each sheet's last cells.
 */
export class SheetSelections {
	private readonly saved = new WeakMap<Worksheet, Selection>();

	remember(sheet: Worksheet | undefined, selection: Selection): void {
		if (sheet) this.saved.set(sheet, selection);
	}

	/** The remembered selection of a sheet, else the one saved in the file, else A1. */
	restore(index: number, workbook: Workbook | undefined): Selection {
		const sheet = workbook?.sheets[index];
		const saved = sheet ? this.saved.get(sheet) : undefined;
		if (!saved) return initialSelection(index, workbook);
		const { drawing: _drawing, ...cells } = saved;
		return { ...cells, sheet: index };
	}
}

export const isWholeRows = (range: CellRange): boolean =>
	range.start.col === 0 && range.end.col === MAX_COL;
export const isWholeCols = (range: CellRange): boolean =>
	range.start.row === 0 && range.end.row === MAX_ROW;
const columnPart = (address: CellAddress): string => formatAddress(address).replace(/\d+$/u, '');

/**
 * One range as Excel writes it: `B2:C5`, whole rows as `2:4` and whole columns as `A:C`. The
 * single formatter for selection-change events, getSelection() and the grid's announcements.
 */
export function rangeRef(range: CellRange): string {
	if (isWholeRows(range) && isWholeCols(range)) return `1:${MAX_ROW + 1}`;
	if (isWholeRows(range)) return `${range.start.row + 1}:${range.end.row + 1}`;
	if (isWholeCols(range)) return `${columnPart(range.start)}:${columnPart(range.end)}`;
	return formatRange(range);
}

/** `B2:C5,E1` for a selection (the order the ranges were made in). */
export function selectionRef(selection: Selection): string {
	return selection.ranges.map(rangeRef).join(',');
}

export function activeRef(selection: Selection): string {
	return formatAddress(selection.active);
}

/**
 * Parses `B2`, `B2:C5`, `A:A`, `1:3`, `B2:C5,E1` and a sheet-qualified `'My Sheet'!B2`. Returns
 * the sheet index (when named) and the ranges, or undefined when the text is not a reference.
 */
export function parseSelectionRef(
	ref: string,
	workbook?: Workbook,
): { sheet?: number; ranges: CellRange[] } | undefined {
	let text = ref.trim();
	let sheet: number | undefined;
	const bang = text.lastIndexOf('!');
	if (bang > 0) {
		const name = text
			.slice(0, bang)
			.replace(/^'(.*)'$/, '$1')
			.replace(/''/g, "'");
		const index =
			workbook?.sheets.findIndex((item) => item.name.toLowerCase() === name.toLowerCase()) ?? -1;
		if (index < 0) return undefined;
		sheet = index;
		text = text.slice(bang + 1);
	}
	const ranges: CellRange[] = [];
	for (const part of text.split(',')) {
		const range = parseRange(part.trim());
		if (!range) return undefined;
		ranges.push(range);
	}
	if (!ranges.length) return undefined;
	return sheet === undefined ? { ranges } : { sheet, ranges };
}
