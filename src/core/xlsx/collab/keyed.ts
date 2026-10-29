// Lists of objects (comments, rules, tables, charts, ...) as keyed entries of a shared map, so
// concurrent edits to different items both survive and an edit to one item replaces only its
// entry. The model has no ids for these items, so identity is tracked per binding: an item keeps
// the key it was written or read with (`PartIds`); a copy of an item (history restores clones) is
// matched by exact content, then by a natural key (a part name, a cell address). An entry is
// `{ i, d }`: `i` the position (ties broken by key) and `d` the plain JSON item.
import type * as Y from 'yjs';
import { isRecord, jsonEqual, setIfChanged } from './schema';

export interface PartSpec<T extends object> {
	/** The plain JSON of an item; undefined leaves the item out of the shared state. */
	encode: (item: T) => unknown;
	/** The item of an entry, or undefined when the entry is malformed. */
	decode: (raw: unknown) => T | undefined;
	/** When set, the entry key is derived from the content (at most one entry per key). */
	fixedKey?: (encoded: unknown) => string;
	/** Matches a copy to its entry when neither identity nor content do. */
	natural?: (encoded: unknown) => string | undefined;
	/** Carries what only the local package knows (a part name) onto a decoded item. */
	adopt?: (decoded: T, local: T | undefined) => void;
}

/** Item identity of one part of one sheet. */
export class PartIds {
	readonly byItem = new WeakMap<object, string>();
	readonly byKey = new Map<string, object>();

	set(item: object, key: string): void {
		this.byItem.set(item, key);
		this.byKey.set(key, item);
	}
}

interface Entry {
	i: number;
	d: unknown;
}

const isEntry = (raw: unknown): raw is Entry =>
	isRecord(raw) && typeof raw.i === 'number' && Number.isFinite(raw.i) && 'd' in raw;

const naturalOf = <T extends object>(spec: PartSpec<T>, encoded: unknown): string | undefined =>
	spec.natural ? spec.natural(encoded) : spec.fixedKey?.(encoded);

function push<K, V>(index: Map<K, V[]>, key: K, value: V): void {
	const list = index.get(key);
	if (list) list.push(value);
	else index.set(key, [value]);
}

/** Writes `items` into `map` where they differ; keys of vanished items are deleted. */
export function writeKeyed<T extends object>(
	map: Y.Map<unknown>,
	items: readonly T[],
	ids: PartIds,
	spec: PartSpec<T>,
	mint: () => string,
): void {
	const existing = new Map<string, unknown>();
	for (const [key, raw] of map.entries()) if (isEntry(raw)) existing.set(key, raw.d);
	const list = items.flatMap((item) => {
		const encoded = spec.encode(item);
		return encoded === undefined ? [] : [{ item, encoded }];
	});
	const keys: (string | undefined)[] = list.map(() => undefined);
	const used = new Set<string>();
	if (spec.fixedKey) {
		list.forEach(({ encoded }, n) => {
			const base = spec.fixedKey?.(encoded) ?? '';
			let key = base;
			for (let dup = 2; used.has(key); dup++) key = `${base}#${dup}`;
			used.add(key);
			keys[n] = key;
		});
	} else {
		list.forEach(({ item }, n) => {
			const key = ids.byItem.get(item);
			if (key !== undefined && existing.has(key) && !used.has(key)) {
				used.add(key);
				keys[n] = key;
			}
		});
		const exact = new Map<string, string[]>();
		const loose = new Map<string, string[]>();
		for (const [key, d] of existing) {
			if (used.has(key)) continue;
			push(exact, JSON.stringify(d), key);
			const nat = naturalOf(spec, d);
			if (nat !== undefined) push(loose, nat, key);
		}
		const take = (index: Map<string, string[]>, at: string | undefined): string | undefined => {
			const found = at === undefined ? undefined : index.get(at);
			while (found?.length) {
				const key = found.shift();
				if (key !== undefined && !used.has(key)) return key;
			}
			return undefined;
		};
		for (const pass of ['exact', 'natural'] as const)
			list.forEach(({ encoded }, n) => {
				if (keys[n] !== undefined) return;
				const key =
					pass === 'exact'
						? take(exact, JSON.stringify(encoded))
						: take(loose, naturalOf(spec, encoded));
				if (key !== undefined) {
					used.add(key);
					keys[n] = key;
				}
			});
		list.forEach((_entry, n) => {
			if (keys[n] !== undefined) return;
			let key = mint();
			while (map.has(key) || used.has(key)) key = mint();
			used.add(key);
			keys[n] = key;
		});
	}
	list.forEach(({ item, encoded }, n) => {
		const key = keys[n];
		if (key === undefined) return;
		ids.set(item, key);
		setIfChanged(map, key, { i: n, d: encoded });
	});
	for (const key of [...map.keys()]) if (!used.has(key)) map.delete(key);
}

/** The items of `map` in order, reusing the objects of `local` that still match; undefined when unchanged. */
export function readKeyed<T extends object>(
	map: Y.Map<unknown>,
	local: readonly T[],
	ids: PartIds,
	spec: PartSpec<T>,
): T[] | undefined {
	const entries = [...map.entries()]
		.filter((pair): pair is [string, Entry] => isEntry(pair[1]))
		.sort((a, b) => a[1].i - b[1].i || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
	const unclaimed = new Set<T>(local);
	const encodedOf = new Map<T, unknown>();
	const encode = (item: T): unknown => {
		if (!encodedOf.has(item)) encodedOf.set(item, spec.encode(item));
		return encodedOf.get(item);
	};
	let exact: Map<string, T[]> | undefined;
	let loose: Map<string, T[]> | undefined;
	const index = (): void => {
		if (exact) return;
		exact = new Map();
		loose = new Map();
		for (const item of local) {
			const encoded = encode(item);
			if (encoded === undefined) continue;
			push(exact, JSON.stringify(encoded), item);
			const nat = naturalOf(spec, encoded);
			if (nat !== undefined) push(loose, nat, item);
		}
	};
	const claim = (found: T[] | undefined): T | undefined => {
		while (found?.length) {
			const item = found.shift();
			if (item !== undefined && unclaimed.has(item)) return item;
		}
		return undefined;
	};
	const next: T[] = [];
	const keep = new Map<string, T>();
	for (const [key, entry] of entries) {
		let counterpart = ids.byKey.get(key) as T | undefined;
		if (counterpart && !unclaimed.has(counterpart)) counterpart = undefined;
		if (!counterpart) {
			index();
			const nat = naturalOf(spec, entry.d);
			counterpart =
				claim(exact?.get(JSON.stringify(entry.d))) ??
				claim(nat === undefined ? undefined : loose?.get(nat));
		}
		let item: T | undefined;
		if (counterpart && jsonEqual(encode(counterpart), entry.d)) item = counterpart;
		else {
			item = spec.decode(entry.d);
			if (!item) continue;
			spec.adopt?.(item, counterpart);
		}
		if (counterpart) unclaimed.delete(counterpart);
		keep.set(key, item);
		next.push(item);
	}
	ids.byKey.clear();
	for (const [key, item] of keep) ids.set(item, key);
	if (next.length === local.length && next.every((item, n) => item === local[n])) return undefined;
	return next;
}
