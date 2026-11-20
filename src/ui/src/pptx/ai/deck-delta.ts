/**
 * What one assistant write changes, recorded against the deck it was computed
 * from so it can be replayed later on whatever the deck has become.
 *
 * Collections of objects with unique string ids (slides, elements, comments,
 * sections) are addressed by id, so an edit still lands on its target after
 * other slides or elements were added, removed or reordered. Everything else is
 * compared by value. A change conflicts when its target no longer holds the
 * value the change was made from.
 */

type Rec = Record<string, unknown>;
type Keyed = Rec & { id: string };

/** A property name, or the id of an item in a keyed collection. */
export type DeckPathStep = string | { id: string };
export type DeckPath = readonly DeckPathStep[];

export type DeckChange =
	| { kind: 'set'; path: DeckPath; before: unknown; after: unknown }
	| { kind: 'insert'; path: DeckPath; item: Keyed; afterId: string | null; atEnd: boolean }
	| { kind: 'remove'; path: DeckPath; item: Keyed }
	| { kind: 'order'; path: DeckPath; before: string[]; after: string[] };

/** Derived from position, so it is recomputed instead of compared. */
const DERIVED_KEYS = new Set(['slideNumber']);

function isRecord(value: unknown): value is Rec {
	if (typeof value !== 'object' || value === null) {
		return false;
	}
	const proto = Object.getPrototypeOf(value) as unknown;
	return proto === Object.prototype || proto === null;
}

function isKeyed(value: unknown): value is Keyed[] {
	if (!Array.isArray(value)) {
		return false;
	}
	const ids = new Set<string>();
	for (const item of value) {
		if (!isRecord(item) || typeof item.id !== 'string' || !item.id || ids.has(item.id)) {
			return false;
		}
		ids.add(item.id);
	}
	return true;
}

function bytesOf(value: unknown): Uint8Array | undefined {
	if (value instanceof ArrayBuffer) {
		return new Uint8Array(value);
	}
	return ArrayBuffer.isView(value)
		? new Uint8Array(value.buffer, value.byteOffset, value.byteLength)
		: undefined;
}

function sameList(a: readonly unknown[], b: readonly unknown[]): boolean {
	return a.length === b.length && a.every((item, index) => sameValue(item, b[index]));
}

/** Deep equality over what a parsed deck holds. Derived keys are ignored. */
export function sameValue(a: unknown, b: unknown): boolean {
	if (Object.is(a, b)) {
		return true;
	}
	if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) {
		return false;
	}
	if (Array.isArray(a) || Array.isArray(b)) {
		return Array.isArray(a) && Array.isArray(b) && sameList(a, b);
	}
	const bytes = bytesOf(a);
	if (bytes) {
		const other = bytesOf(b);
		return !!other && other.length === bytes.length && bytes.every((v, i) => v === other[i]);
	}
	if (a instanceof Date || b instanceof Date) {
		return a instanceof Date && b instanceof Date && a.getTime() === b.getTime();
	}
	if ((a instanceof Map && b instanceof Map) || (a instanceof Set && b instanceof Set)) {
		return sameList([...a], [...b]);
	}
	if (!isRecord(a) || !isRecord(b)) {
		return false;
	}
	const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
	return [...keys].every((key) => DERIVED_KEYS.has(key) || sameValue(a[key], b[key]));
}

function diffCollection(a: Keyed[], b: Keyed[], path: DeckPath, out: DeckChange[]): void {
	const before = new Map(a.map((item) => [item.id, item]));
	const after = new Set(b.map((item) => item.id));
	for (const item of a) {
		if (!after.has(item.id)) {
			out.push({ kind: 'remove', path, item });
		}
	}
	const keptBefore = a.filter((item) => after.has(item.id)).map((item) => item.id);
	const keptAfter = b.filter((item) => before.has(item.id)).map((item) => item.id);
	if (!sameList(keptBefore, keptAfter)) {
		out.push({ kind: 'order', path, before: keptBefore, after: keptAfter });
	}
	const lastKept = b.reduce((last, item, index) => (before.has(item.id) ? index : last), -1);
	b.forEach((item, index) => {
		const previous = before.get(item.id);
		if (previous) {
			diff(previous, item, [...path, { id: item.id }], out);
		} else {
			const afterId = b[index - 1]?.id ?? null;
			out.push({ kind: 'insert', path, item, afterId, atEnd: index > lastKept });
		}
	});
}

function diff(a: unknown, b: unknown, path: DeckPath, out: DeckChange[]): void {
	if (isRecord(a) && isRecord(b)) {
		for (const key of new Set([...Object.keys(a), ...Object.keys(b)])) {
			if (!DERIVED_KEYS.has(key)) {
				diff(a[key], b[key], [...path, key], out);
			}
		}
	} else if (isKeyed(a) && isKeyed(b) && a.length + b.length > 0) {
		diffCollection(a, b, path, out);
	} else if (!sameValue(a, b)) {
		out.push({ kind: 'set', path, before: a, after: b });
	}
}

/** The changes that turn `before` into `after`. */
export function diffDeck(before: object, after: object): DeckChange[] {
	const out: DeckChange[] = [];
	diff(before, after, [], out);
	return out;
}

const MISSING = Symbol('missing');

function read(root: unknown, path: DeckPath): unknown {
	let node = root;
	for (const step of path) {
		if (typeof step === 'string') {
			if (!isRecord(node)) {
				return MISSING;
			}
			node = node[step];
		} else {
			node = Array.isArray(node)
				? (node.find((item) => isRecord(item) && item.id === step.id) ?? MISSING)
				: MISSING;
		}
		if (node === MISSING) {
			return MISSING;
		}
	}
	return node;
}

/** A readable location for messages: `slides/slide-1/elements/el-2/text`. */
export function describeDeckPath(path: DeckPath): string {
	return path.map((step) => (typeof step === 'string' ? step : step.id)).join('/');
}

function conflictOf(change: DeckChange, root: unknown): boolean {
	if (change.kind === 'set') {
		const current = read(root, change.path);
		return (
			current === MISSING ||
			(!sameValue(current, change.before) && !sameValue(current, change.after))
		);
	}
	const list = read(root, change.path);
	if (!isKeyed(list)) {
		return true;
	}
	if (change.kind === 'order') {
		const present = new Set(list.map((item) => item.id));
		const moved = new Set(change.before);
		return !sameList(
			list.filter((item) => moved.has(item.id)).map((item) => item.id),
			change.before.filter((id) => present.has(id)),
		);
	}
	const current = list.find((item) => item.id === change.item.id);
	return current !== undefined && !sameValue(current, change.item);
}

/** The first change that no longer fits `root`, or `undefined` when all of them apply. */
export function findDeckConflict(
	changes: readonly DeckChange[],
	root: object,
): DeckChange | undefined {
	return changes.find((change) => conflictOf(change, root));
}

function insertItem(list: Keyed[], change: Extract<DeckChange, { kind: 'insert' }>): void {
	if (list.some((item) => item.id === change.item.id)) {
		return;
	}
	const anchor = list.findIndex((item) => item.id === change.afterId);
	if (change.atEnd || (change.afterId !== null && anchor < 0)) {
		list.push(change.item);
	} else {
		list.splice(anchor + 1, 0, change.item);
	}
}

function reorder(list: Keyed[], ids: readonly string[]): void {
	const moved = new Set(ids);
	const byId = new Map(list.map((item) => [item.id, item]));
	const queue = ids.flatMap((id) => byId.get(id) ?? []);
	list.forEach((item, index) => {
		if (moved.has(item.id)) {
			list[index] = queue.shift() as Keyed;
		}
	});
}

function applyChange(change: DeckChange, root: unknown): void {
	if (change.kind === 'set') {
		const parent = read(root, change.path.slice(0, -1));
		const key = change.path.at(-1);
		if (isRecord(parent) && typeof key === 'string') {
			if (change.after === undefined) {
				delete parent[key];
			} else {
				parent[key] = change.after;
			}
		}
		return;
	}
	const list = read(root, change.path) as Keyed[];
	if (change.kind === 'insert') {
		insertItem(list, change);
	} else if (change.kind === 'order') {
		reorder(list, change.after);
	} else {
		const index = list.findIndex((item) => item.id === change.item.id);
		if (index >= 0) {
			list.splice(index, 1);
		}
	}
}

/** Replays `changes` on `root` in place. Check {@link findDeckConflict} first. */
export function applyDeckChanges(changes: readonly DeckChange[], root: object): void {
	for (const change of changes) {
		applyChange(structuredClone(change), root);
	}
	const slides = (root as { slides?: { slideNumber?: number }[] }).slides;
	const moved = changes.some(
		(c) => c.kind !== 'set' && c.path.length === 1 && c.path[0] === 'slides',
	);
	if (moved && slides) {
		slides.forEach((slide, index) => {
			slide.slideNumber = index + 1;
		});
	}
}
