// The user's own sort custom lists (Excel's Options > Edit Custom Lists, application wide).
// Kept in the browser's localStorage when it is available and for the session otherwise.
import { getCell, isCellError, type CellRange, type Worksheet } from 'ooxml-core/xlsx';

const KEY = 'ooxml.xlsx.customLists';
let memory: string[][] = [];

function write(lists: string[][]): void {
	memory = lists.map((list) => [...list]);
	try {
		globalThis.localStorage?.setItem(KEY, JSON.stringify(memory));
	} catch {
		// Storage can be blocked; the list then lasts for the session.
	}
}

function read(): string[][] {
	try {
		const raw = globalThis.localStorage?.getItem(KEY);
		if (raw === null || raw === undefined) return memory;
		const parsed: unknown = JSON.parse(raw);
		if (!Array.isArray(parsed)) return memory;
		return parsed
			.filter(
				(list): list is string[] => Array.isArray(list) && list.every((s) => typeof s === 'string'),
			)
			.filter((list) => list.length > 0);
	} catch {
		return memory;
	}
}

const cleanItems = (items: readonly string[]): string[] =>
	items.map((item) => item.trim()).filter((item) => item.length > 0);

/** The saved user lists, oldest first. */
export const userSortLists = (): string[][] => read().map((list) => [...list]);

/** Saves a list unless an equal one exists; returns its index among the user lists. */
export function saveUserSortList(items: readonly string[]): number {
	const cleaned = cleanItems(items);
	if (cleaned.length === 0) return -1;
	const lists = read();
	const same = lists.findIndex(
		(list) => list.length === cleaned.length && list.every((item, i) => item === cleaned[i]),
	);
	if (same >= 0) return same;
	write([...lists, cleaned]);
	return memory.length - 1;
}

/** Replaces the user list at `index` with `items` (rename/edit); returns false if invalid. */
export function updateUserSortList(index: number, items: readonly string[]): boolean {
	const cleaned = cleanItems(items);
	const lists = read();
	if (index < 0 || index >= lists.length || cleaned.length === 0) return false;
	const next = lists.map((list, i) => (i === index ? cleaned : [...list]));
	write(next);
	return true;
}

/** Deletes the user list at `index`; returns false if out of range. */
export function deleteUserSortList(index: number): boolean {
	const lists = read();
	if (index < 0 || index >= lists.length) return false;
	write(lists.filter((_, i) => i !== index));
	return true;
}

/** Extracts non-empty cell values from `range` in row-major order for Custom Lists import. */
export function importSortListFromRange(sheet: Worksheet, range: CellRange): string[] {
	const r0 = Math.min(range.start.row, range.end.row);
	const r1 = Math.max(range.start.row, range.end.row);
	const c0 = Math.min(range.start.col, range.end.col);
	const c1 = Math.max(range.start.col, range.end.col);
	const out: string[] = [];
	for (let r = r0; r <= r1; r++) {
		for (let c = c0; c <= c1; c++) {
			const cell = getCell(sheet, r, c);
			if (!cell || cell.value === null || isCellError(cell.value)) continue;
			const text = String(cell.value).trim();
			if (text.length > 0) out.push(text);
		}
	}
	return out;
}

/** Forgets every user list (tests, and a reset). */
export function clearUserSortLists(): void {
	memory = [];
	try {
		globalThis.localStorage?.removeItem(KEY);
	} catch {
		// Nothing to remove.
	}
}
