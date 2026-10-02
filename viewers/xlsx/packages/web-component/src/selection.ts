/** The selection model behind `ctx.selection`, and A1 text conversions for the element API. */
import {
	formatAddress,
	formatRange,
	normalizeRange,
	parseRange,
	type CellAddress,
	type CellRange,
	type Workbook,
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

/** `B2:C5,E1` for a selection (the order the ranges were made in). */
export function selectionRef(selection: Selection): string {
	return selection.ranges.map(formatRange).join(',');
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
