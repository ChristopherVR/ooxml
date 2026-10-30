import type { VisioDataColumn, VisioDataRecordset } from './data-recordsets';

/** How the rows of a refreshed table line up with the rows the drawing was linked to. */
export interface VisioRowMatch {
	/** The new row index for each old row index, or undefined when no new row matches it. */
	map: (number | undefined)[];
	/** The key column the match used; undefined means rows were matched by position. */
	key?: string;
	/** The key column is gone from the refreshed table, so no old row can be matched. */
	keyMissing: boolean;
}

/**
 * Match old rows to refreshed rows. With a primary key the key column's text decides (the first
 * new row wins a duplicate key); without one the row position does, as Visio does for a
 * recordset that has no primary key.
 */
export function matchRefreshedRows(
	oldColumns: readonly string[],
	oldRows: readonly (readonly string[])[],
	newColumns: readonly string[],
	newRows: readonly (readonly string[])[],
	key?: string,
): VisioRowMatch {
	if (!key)
		return {
			map: oldRows.map((_, index) => (index < newRows.length ? index : undefined)),
			keyMissing: false,
		};
	const oldIndex = oldColumns.indexOf(key);
	const newIndex = newColumns.indexOf(key);
	if (oldIndex < 0 || newIndex < 0)
		return { map: oldRows.map(() => undefined), key, keyMissing: true };
	const byKey = new Map<string, number>();
	newRows.forEach((row, index) => {
		const value = (row[newIndex] ?? '').trim();
		if (value !== '' && !byKey.has(value)) byKey.set(value, index);
	});
	return {
		map: oldRows.map((row) => byKey.get((row[oldIndex] ?? '').trim())),
		key,
		keyMissing: false,
	};
}

/** What a refresh will do to the shapes linked to a recordset, for the message the user reads. */
export interface VisioRefreshReport {
	key?: string;
	keyMissing: boolean;
	/** Linked shapes whose row was found and whose values will update. */
	matched: number;
	/** Linked shapes whose row no longer exists; they keep their values and are unlinked. */
	unmatched: { shapeId: string; pageId: string; rowId: string }[];
	/** Refreshed rows that no old row matched. */
	added: number;
}

export function visioRefreshReport(
	set: VisioDataRecordset,
	columns: readonly VisioDataColumn[],
	rows: readonly (readonly string[])[],
): VisioRefreshReport {
	const match = matchRefreshedRows(
		set.columns.map((column) => column.name),
		set.rows.map((row) => row.values),
		columns.map((column) => column.name),
		rows,
		set.primaryKey,
	);
	const rowIndex = new Map(set.rows.map((row, index) => [row.id, index]));
	const used = new Set<number>();
	const unmatched: VisioRefreshReport['unmatched'] = [];
	let matched = 0;
	for (const link of set.links) {
		const target = match.map[rowIndex.get(link.rowId) ?? -1];
		if (target === undefined) unmatched.push(link);
		else matched++;
	}
	for (const target of match.map) if (target !== undefined) used.add(target);
	return {
		...(match.key ? { key: match.key } : {}),
		keyMissing: match.keyMissing,
		matched,
		unmatched,
		added: rows.length - used.size,
	};
}
