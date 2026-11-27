/**
 * table-style-editor.ts - pure, runtime-agnostic helpers for the parts of
 * "edit tableStyles.xml" that are not a per-section XML merge:
 * creating a brand-new style entry (with a fresh GUID), and removing one
 * from a {@link ParsedTableStyleMap}.
 *
 * These operate on the TYPED model only (`ParsedTableStyleMap` /
 * `ParsedTableStyleEntry`), the same object shape `PptxData.tableStyleMap`
 * already exposes and `PptxSaveOptions.tableStyles` already accepts. A
 * caller (SDK consumer, or the `create_table_style` / `delete_table_style`
 * MCP tools) mutates the loaded map with these helpers, then passes it -
 * along with `tableStylesDefaultId` / `tableStylesToDelete` for the two
 * archive-level operations these helpers do NOT perform - to
 * `handler.save(slides, { tableStyles, tableStylesDefaultId,
 * tableStylesToDelete })`, where `applyTableStylesPart`
 * (`PptxHandlerRuntimeSaveViewProperties.ts`) does the actual
 * `ppt/tableStyles.xml` create/delete/re-default.
 *
 * @module table-style-editor
 */
import { createOfficeGuid } from '../../../../crypto/uuid';
import { randomIdSource, type IdSource } from '../../../../opc/reproducible';
import type {
	CreateTableStyleOptions,
	ParsedTableStyleEntry,
	ParsedTableStyleMap,
} from '../../types';
import { normalizeTableStyleGuid } from './table-style-entry-parse';

/**
 * Deep-clone a plain JSON-shaped value (no functions, dates, or cycles - true
 * of every field on {@link ParsedTableStyleEntry}). Prefers the platform
 * `structuredClone` (Node 17+/Bun/browsers) and falls back to a JSON
 * round-trip for older runtimes.
 */
function deepClone<T>(value: T): T {
	const clone = (globalThis as { structuredClone?: <V>(v: V) => V }).structuredClone;
	if (clone) {
		return clone(value);
	}
	return JSON.parse(JSON.stringify(value)) as T;
}

/**
 * Generate a fresh, braced-and-uppercased GUID for a new table style, unique
 * within `existingIds`, from `ids` (random by default; the shared
 * `createOfficeGuid`).
 */
export function generateTableStyleGuid(
	existingIds: ReadonlySet<string> = new Set(),
	ids: IdSource = randomIdSource,
): string {
	for (;;) {
		const guid = createOfficeGuid(ids);
		if (!existingIds.has(normalizeTableStyleGuid(guid))) {
			return guid;
		}
	}
}

/**
 * Build a brand-new {@link ParsedTableStyleEntry}: a fresh GUID (or the
 * caller's explicit one), the given display name, and, when {@link
 * CreateTableStyleOptions.basedOn} is supplied, a deep clone of every
 * fill/text/border/cell3D/background section from that entry as the starting
 * point. Does NOT mutate any map; combine with {@link addTableStyleToMap}.
 */
export function createTableStyleEntry(
	existingMap: ParsedTableStyleMap,
	options: CreateTableStyleOptions,
): ParsedTableStyleEntry {
	const styleId = options.styleId
		? normalizeTableStyleGuid(options.styleId)
		: generateTableStyleGuid(new Set(Object.keys(existingMap).map(normalizeTableStyleGuid)));

	if (!options.basedOn) {
		return { styleId, styleName: options.styleName };
	}

	const cloned = deepClone(options.basedOn);
	return { ...cloned, styleId, styleName: options.styleName };
}

/**
 * Add (or replace) an entry on a {@link ParsedTableStyleMap} in place, keyed
 * by its normalised `styleId`. Returns the same map for chaining.
 */
export function addTableStyleToMap(
	map: ParsedTableStyleMap,
	entry: ParsedTableStyleEntry,
): ParsedTableStyleMap {
	map[normalizeTableStyleGuid(entry.styleId)] = entry;
	return map;
}

/**
 * Remove a style from a {@link ParsedTableStyleMap} in place. Returns whether
 * an entry was actually present (and removed). This only edits the in-memory
 * map the way any other section edit does; the caller must ALSO pass the
 * GUID via `PptxSaveOptions.tableStylesToDelete` at save time for the removal
 * to reach `ppt/tableStyles.xml` (map omission alone is deliberately treated
 * as "untouched", not "deleted" - see `PptxSaveOptions.tableStylesToDelete`'s
 * docs for why).
 */
export function deleteTableStyleFromMap(map: ParsedTableStyleMap, styleId: string): boolean {
	const normalized = normalizeTableStyleGuid(styleId);
	if (!(normalized in map)) {
		return false;
	}
	delete map[normalized];
	return true;
}
