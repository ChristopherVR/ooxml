import type { Workbook, Worksheet } from '../model';
import type { SheetMatch, WorkbookDifference } from './types';

/** Two sheets taken to be the same sheet, by their zero-based positions. */
export interface SheetPair {
	before: number;
	after: number;
	matchedBy: 'name' | 'sheetId';
}

/**
 * Pairs the sheets of two workbooks: by name (case-insensitive, as Excel compares sheet names),
 * then by `sheetId` among the sheets left over, which finds renamed sheets. Pairs are in the
 * after workbook's order.
 */
export function matchSheets(before: Workbook, after: Workbook, mode: SheetMatch): SheetPair[] {
	const pairs: SheetPair[] = [];
	const pairedBefore = new Set<number>();
	const pairedAfter = new Set<number>();
	/** Pairs each sheet left in `before` with the sheet left in `after` that has the same key. */
	const pairBy = <K>(key: (sheet: Worksheet) => K, matchedBy: SheetPair['matchedBy']) => {
		const unpaired = new Map<K, number>();
		after.sheets.forEach((sheet, index) => {
			if (!pairedAfter.has(index)) unpaired.set(key(sheet), index);
		});
		before.sheets.forEach((sheet, index) => {
			if (pairedBefore.has(index)) return;
			const match = unpaired.get(key(sheet));
			if (match === undefined || pairedAfter.has(match)) return;
			pairs.push({ before: index, after: match, matchedBy });
			pairedBefore.add(index);
			pairedAfter.add(match);
		});
	};
	if (mode !== 'sheetId') pairBy((sheet) => sheet.name.toLowerCase(), 'name');
	if (mode !== 'name') pairBy((sheet) => sheet.sheetId, 'sheetId');
	return pairs.sort((x, y) => x.after - y.after);
}

/**
 * Positions (into `values`) of one longest strictly increasing subsequence: the items (sheets,
 * conditional format rules) that kept their relative order. Every other item moved.
 */
export function longestIncreasing(values: readonly number[]): Set<number> {
	const tails: number[] = [];
	const previous: number[] = [];
	values.forEach((value, i) => {
		let lo = 0;
		let hi = tails.length;
		while (lo < hi) {
			const mid = (lo + hi) >> 1;
			if ((values[tails[mid] ?? 0] ?? 0) < value) lo = mid + 1;
			else hi = mid;
		}
		previous[i] = tails[lo - 1] ?? -1;
		tails[lo] = i;
	});
	const keep = new Set<number>();
	for (let i = tails.at(-1) ?? -1; i >= 0; i = previous[i] ?? -1) keep.add(i);
	return keep;
}

/** Sheets removed, added, renamed, moved and shown or hidden. */
export function sheetDifferences(
	before: Workbook,
	after: Workbook,
	pairs: readonly SheetPair[],
): WorkbookDifference[] {
	const out: WorkbookDifference[] = [];
	const pairedBefore = new Set(pairs.map((p) => p.before));
	const pairedAfter = new Set(pairs.map((p) => p.after));
	before.sheets.forEach((sheet, index) => {
		if (!pairedBefore.has(index)) out.push({ kind: 'sheetRemoved', sheet: sheet.name, index });
	});
	after.sheets.forEach((sheet, index) => {
		if (!pairedAfter.has(index)) out.push({ kind: 'sheetAdded', sheet: sheet.name, index });
	});
	const byBefore = [...pairs].sort((x, y) => x.before - y.before);
	const kept = longestIncreasing(byBefore.map((p) => p.after));
	const moved = new Set(byBefore.filter((_, i) => !kept.has(i)));
	for (const p of pairs) {
		const b = before.sheets[p.before];
		const a = after.sheets[p.after];
		if (!b || !a) continue;
		if (b.name !== a.name) out.push({ kind: 'sheetRenamed', from: b.name, to: a.name });
		if (moved.has(p)) out.push({ kind: 'sheetMoved', sheet: a.name, from: p.before, to: p.after });
		if (b.state !== a.state)
			out.push({ kind: 'sheetState', sheet: a.name, from: b.state, to: a.state });
	}
	return out;
}
