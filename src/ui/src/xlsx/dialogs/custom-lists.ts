// The user's own sort custom lists (Excel's Options > Edit Custom Lists, application wide).
// Kept in the browser's localStorage when it is available and for the session otherwise.
const KEY = 'ooxml.xlsx.customLists';
let memory: string[][] = [];

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

/** The saved user lists, oldest first. */
export const userSortLists = (): string[][] => read().map((list) => [...list]);

/** Saves a list unless an equal one exists; returns its index among the user lists. */
export function saveUserSortList(items: readonly string[]): number {
	const lists = read();
	const same = lists.findIndex(
		(list) => list.length === items.length && list.every((item, i) => item === items[i]),
	);
	if (same >= 0) return same;
	memory = [...lists, [...items]];
	try {
		globalThis.localStorage?.setItem(KEY, JSON.stringify(memory));
	} catch {
		// Storage can be blocked; the list then lasts for the session.
	}
	return memory.length - 1;
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
