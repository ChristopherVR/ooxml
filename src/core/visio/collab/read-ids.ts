// Unique ids for the items two peers numbered alike (pages, masters, shapes). Every peer
// computes the same result from the room alone:
//   - an item whose renumbering is already persisted (`visio:renum`, written by a fix-up, see
//     fixups.ts) keeps its persisted id and is visited first, so a later concurrent add cannot
//     displace it;
//   - the others are visited in key order; one whose id is taken gets the next number above the
//     largest id in use and is reported as freshly renumbered, which is what a fix-up persists;
//   - the item that kept an id other items were moved off is reported as a claim, and the fix-up
//     persists it too (`i: [id, id]`), so an item added later, concurrently, with a lower key cannot
//     displace an id peers have already seen persisted.
import type { RenumEntry, SharedDrawing } from './schema';
import { renumEntry } from './schema';

export interface IdItem {
	key: string;
	/** The id the author gave the item. */
	id: string;
	/** The persisted id, when a renumbering was persisted for exactly this author id. */
	fixed?: string | undefined;
}
export interface Renumbered {
	from: string;
	to: string;
}

export const compare = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);
export const numeric = (id: string): number => (/^\d+$/.test(id) ? Number(id) : 0);

export function assignIds(items: readonly IdItem[]): {
	ids: Map<string, string>;
	fresh: Map<string, Renumbered>;
	/** Items that keep an id another item had to leave and that is not persisted yet. */
	claims: Set<string>;
} {
	const effective = (item: IdItem): string => item.fixed ?? item.id;
	let max = Math.max(0, ...items.map((item) => numeric(effective(item))));
	const order = [...items].sort(
		(a, b) =>
			Number(b.fixed !== undefined) - Number(a.fixed !== undefined) || compare(a.key, b.key),
	);
	const ids = new Map<string, string>();
	const fresh = new Map<string, Renumbered>();
	const holder = new Map<string, IdItem>();
	const claims = new Set<string>();
	for (const item of order) {
		let id = effective(item);
		const taken = holder.get(id);
		if (taken) {
			if (taken.fixed === undefined) claims.add(taken.key);
			id = String(++max);
			fresh.set(item.key, { from: item.id, to: id });
		}
		holder.set(id, item);
		ids.set(item.key, id);
	}
	return { ids, fresh, claims };
}

/** The persisted renumbering of `key`, when it applies to an item the author numbered `id`. */
export const fixedId = (
	renum: ReadonlyMap<string, RenumEntry>,
	key: string,
	id: string,
): string | undefined => {
	const pair = renum.get(key)?.i;
	return pair && pair[0] === id ? pair[1] : undefined;
};

/** Every renumbering record in the room by key. */
export function readRenum(shared: SharedDrawing): Map<string, RenumEntry> {
	const out = new Map<string, RenumEntry>();
	for (const [key, value] of shared.renum) {
		const entry = renumEntry(value);
		if (entry) out.set(key, entry);
	}
	return out;
}

/** The renumbering record of `key` with the change or claim of this read applied. */
export function renumRecord(
	renum: ReadonlyMap<string, RenumEntry>,
	recordKey: string,
	item: { id: string },
	change: Renumbered | undefined,
	claimed: boolean,
): RenumEntry | undefined {
	if (change) return { ...renum.get(recordKey), i: [change.from, change.to] };
	if (claimed) return { ...renum.get(recordKey), i: [item.id, item.id] };
	return undefined;
}

/** A write a peer that can write makes to persist what every peer computes on read. */
export interface Fixup {
	map: 'renum';
	key: string;
	value: unknown;
}
