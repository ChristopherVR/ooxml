// The shared workbook schema. Four top-level Yjs maps (top-level types never conflict when two
// peers initialise a room at once, unlike nested maps created concurrently):
//   xlsx:meta    { v: schema version, defaultStyle: style key, date1904 }
//   xlsx:styles  style key -> resolved CellStyle (content-addressed, see styles.ts)
//   xlsx:names   `${sheetKey}!${lower-case name}` -> NameEntry
//   xlsx:sheets  sheet key -> Y.Map { props, cells, rows, cols, merges }
// A sheet key is a stable id minted by the peer that added the sheet (never its name or index).
// Inside a sheet: `props` (Y.Map: name, sheetId, order, state, tab, rh, cw), `cells` (`A1` ->
// CellEntry), `rows` (1-based row number -> RowEntry), `cols` (`C:E` span -> ColEntry) and
// `merges` (`A1:B2` -> true). Entries are plain JSON, so a cell is replaced atomically.
import * as Y from 'yjs';

export const SCHEMA_VERSION = 1;

export const META = 'xlsx:meta';
export const STYLES = 'xlsx:styles';
export const NAMES = 'xlsx:names';
export const SHEETS = 'xlsx:sheets';

export const SHEET_PARTS = ['props', 'cells', 'rows', 'cols', 'merges'] as const;
export type SheetPart = (typeof SHEET_PARTS)[number];

export interface SharedWorkbook {
	meta: Y.Map<unknown>;
	styles: Y.Map<unknown>;
	names: Y.Map<unknown>;
	sheets: Y.Map<unknown>;
}

/** The workbook's top-level maps in `doc`. */
export function sharedWorkbook(doc: Y.Doc): SharedWorkbook {
	return {
		meta: doc.getMap<unknown>(META),
		styles: doc.getMap<unknown>(STYLES),
		names: doc.getMap<unknown>(NAMES),
		sheets: doc.getMap<unknown>(SHEETS),
	};
}

/** Every top-level map, for an UndoManager scope or deep observers. */
export const sharedTypes = (shared: SharedWorkbook): Y.Map<unknown>[] => [
	shared.meta,
	shared.styles,
	shared.names,
	shared.sheets,
];

/** A nested map, or undefined when the entry is missing or malformed. */
export function childMap(parent: Y.Map<unknown>, key: string): Y.Map<unknown> | undefined {
	const value = parent.get(key);
	return value instanceof Y.Map ? (value as Y.Map<unknown>) : undefined;
}

/** A sheet's part map, created when missing (only ever by a writer, inside a transaction). */
export function sheetPart(sheet: Y.Map<unknown>, part: SheetPart): Y.Map<unknown> {
	const existing = childMap(sheet, part);
	if (existing) return existing;
	const created = new Y.Map<unknown>();
	sheet.set(part, created);
	return created;
}

/** A new sheet map with all its parts, so no other peer ever has to create them. */
export function newSheetMap(): Y.Map<unknown> {
	const sheet = new Y.Map<unknown>();
	for (const part of SHEET_PARTS) sheet.set(part, new Y.Map<unknown>());
	return sheet;
}

/** Deep equality of JSON entries (key order ignored). */
export function jsonEqual(a: unknown, b: unknown): boolean {
	if (a === b) return true;
	if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) return false;
	if (Array.isArray(a) !== Array.isArray(b)) return false;
	const ka = Object.keys(a).filter((k) => (a as Record<string, unknown>)[k] !== undefined);
	const kb = Object.keys(b).filter((k) => (b as Record<string, unknown>)[k] !== undefined);
	if (ka.length !== kb.length) return false;
	return ka.every((k) =>
		jsonEqual((a as Record<string, unknown>)[k], (b as Record<string, unknown>)[k]),
	);
}

/** Sets `key` only when its JSON value differs; returns whether it wrote. */
export function setIfChanged(map: Y.Map<unknown>, key: string, value: unknown): boolean {
	if (value === undefined) {
		if (!map.has(key)) return false;
		map.delete(key);
		return true;
	}
	if (jsonEqual(map.get(key), value)) return false;
	map.set(key, value);
	return true;
}

export const isRecord = (value: unknown): value is Record<string, unknown> =>
	typeof value === 'object' && value !== null && !Array.isArray(value);
