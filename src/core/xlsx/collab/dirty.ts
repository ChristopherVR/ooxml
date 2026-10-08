// Collects what a remote transaction touched from Yjs deep events, so applying it to the model
// visits only the changed cells, rows and sheets. Events must be read inside the observer.
import type * as Y from 'yjs';

export interface SheetDirty {
	full: boolean;
	rows: Set<string>;
	cols: boolean;
	merges: boolean;
	cells: Set<string>;
}

export interface Dirty {
	/** Re-read everything (the workbook settings changed, or a room is adopted). */
	all: boolean;
	/** Sheets were added, removed, renamed, moved or changed properties. */
	list: boolean;
	names: boolean;
	sheets: Map<string, SheetDirty>;
}

export const emptyDirty = (all = false): Dirty => ({
	all,
	list: false,
	names: false,
	sheets: new Map(),
});

export const isDirty = (dirty: Dirty): boolean =>
	dirty.all || dirty.list || dirty.names || dirty.sheets.size > 0;

function sheetDirty(dirty: Dirty, key: string): SheetDirty {
	let entry = dirty.sheets.get(key);
	if (!entry) {
		entry = { full: false, rows: new Set(), cols: false, merges: false, cells: new Set() };
		dirty.sheets.set(key, entry);
	}
	return entry;
}

type AnyEvent = Y.YEvent<Y.AbstractType<unknown>>;

const changedKeys = (event: AnyEvent): string[] => [...event.changes.keys.keys()];

/** Records the events of one observed top-level map. */
export function collectDirty(
	dirty: Dirty,
	top: 'meta' | 'styles' | 'names' | 'sheets',
	events: AnyEvent[],
): void {
	if (top === 'styles') return; // Formats are content-addressed and read on demand.
	if (top === 'meta') {
		dirty.all = true;
		return;
	}
	if (top === 'names') {
		dirty.names = true;
		return;
	}
	for (const event of events) {
		const path = event.path;
		const key = path[0];
		const part = path[1];
		if (key === undefined) {
			dirty.list = true;
			for (const changed of changedKeys(event)) sheetDirty(dirty, changed).full = true;
			continue;
		}
		const entry = sheetDirty(dirty, String(key));
		if (part === undefined || path.length > 2) entry.full = true;
		else if (part === 'props') dirty.list = true;
		else if (part === 'cells') for (const k of changedKeys(event)) entry.cells.add(k);
		else if (part === 'rows') for (const k of changedKeys(event)) entry.rows.add(k);
		else if (part === 'cols') entry.cols = true;
		else if (part === 'merges') entry.merges = true;
		else entry.full = true;
	}
}
