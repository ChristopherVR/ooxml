// Defined names in a shared workbook: keyed by scope (the sheet key, empty for workbook scope)
// and the lower-case name, so the same name added twice converges to one entry.
import type * as Y from 'yjs';
import type { DefinedName } from '../model';
import { isRecord, setIfChanged } from './schema';

export interface NameEntry {
	n: string;
	f: string;
	/** Sheet key of a sheet-scoped name. */
	k?: string;
	h?: 1;
	c?: string;
}

export const nameKey = (name: string, sheetKey: string | undefined): string =>
	`${sheetKey ?? ''}!${name.toLowerCase()}`;

/** Writes every defined name; names whose sheet is not shared are skipped. */
export function writeNames(
	names: Y.Map<unknown>,
	list: readonly DefinedName[],
	sheetKeyAt: (index: number) => string | undefined,
): void {
	const keys = new Set<string>();
	for (const name of list) {
		const scope = name.localSheet === undefined ? undefined : sheetKeyAt(name.localSheet);
		if (name.localSheet !== undefined && scope === undefined) continue;
		const entry: NameEntry = { n: name.name, f: name.formula };
		if (scope !== undefined) entry.k = scope;
		if (name.hidden) entry.h = 1;
		if (name.comment !== undefined) entry.c = name.comment;
		const key = nameKey(name.name, scope);
		keys.add(key);
		setIfChanged(names, key, entry);
	}
	for (const key of [...names.keys()]) if (!keys.has(key)) names.delete(key);
}

/** Reads the defined names; names scoped to a sheet that no longer exists are dropped. */
export function readNames(
	names: Y.Map<unknown>,
	sheetIndexOf: (key: string) => number | undefined,
): DefinedName[] {
	const out: DefinedName[] = [];
	for (const key of [...names.keys()].sort()) {
		const raw = names.get(key);
		if (!isRecord(raw) || typeof raw.n !== 'string' || typeof raw.f !== 'string') continue;
		if (!raw.n || raw.n.length > 255 || raw.f.length > 8192) continue;
		const name: DefinedName = { name: raw.n, formula: raw.f };
		if (typeof raw.k === 'string') {
			const index = sheetIndexOf(raw.k);
			if (index === undefined) continue;
			name.localSheet = index;
		}
		if (raw.h === 1) name.hidden = true;
		if (typeof raw.c === 'string') name.comment = raw.c;
		out.push(name);
	}
	return out;
}
