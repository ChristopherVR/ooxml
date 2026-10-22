// Resolves the shapes of one page from the shared maps: tombstoned and orphaned shapes dropped,
// siblings in their shared order, ids that two peers minted concurrently made unique, and the XML
// of each top-level shape rebuilt from its shell and cell entries in a fixed order.
//
// Id rule (deterministic, so every peer renumbers alike, see read-ids.ts): a shape whose renumbering
// is persisted keeps it; the others are visited in key order, and one whose id is already taken
// takes the next number above the page's largest id and keeps its key. A peer that can write then
// persists the result (fixups.ts), so the id no longer depends on what is added later.
// `Sheet.<old>` in the F attribute of a cell is rewritten to the new id in (a) the renumbered
// shape and every shape nested in it, and (b) every shape written in the same batch as the
// renumbered one (`b`), whose author numbered them together. It is not rewritten in other
// shapes: a reference written after its author had adopted the new id cannot be told from a
// reference to the shape that kept the old id. Masters renumbered the same way rewrite `Master`
// on the shapes of their batch.
import { buildShape, type PieceEntry } from './pieces';
import { type Fixup, assignIds, compare, fixedId } from './read-ids';
import { attrOf } from './scan';
import { type RenumEntry, type ShapeEntry } from './schema';

export interface ResolvedShape {
	key: string;
	/** Final id, after renumbering. */
	id: string;
	parent?: string;
	open: string;
	entries: Map<string, PieceEntry>;
}

export function setAttr(open: string, name: string, value: string): string {
	return open.replace(
		new RegExp(String.raw`(\s${name}\s*=\s*)(?:"[^"]*"|'[^']*')`),
		(_m, head: string) => `${head}"${value}"`,
	);
}

/** Rewrites `Sheet.<n>` for the ids in `remap` inside the F attributes of one cell element. */
export function renumberFormulas(xml: string, remap: ReadonlyMap<string, string>): string {
	if (!remap.size || !xml.includes('Sheet.')) return xml;
	return xml.replace(/(\sF\s*=\s*)("[^"]*"|'[^']*')/g, (_m, head: string, value: string) => {
		const body = value.slice(1, -1).replace(/Sheet\.(\d+)(?!\d)/g, (m, id: string) => {
			const next = remap.get(id);
			return next === undefined ? m : `Sheet.${next}`;
		});
		return `${head}${value[0]}${body}${value[0]}`;
	});
}

/** Rewrites the ids in `remap` on `Shape` tags and in `Sheet.<n>` references of one shape. */
export function renumberShape(xml: string, remap: ReadonlyMap<string, string>): string {
	if (!remap.size) return xml;
	return xml
		.replace(/<(?:\w+:)?Shape\s[^>]*>/g, (tag) =>
			tag.replace(/(\bID\s*=\s*)(["'])(\d+)\2/, (m, head: string, q: string, id: string) => {
				const next = remap.get(id);
				return next === undefined ? m : `${head}${q}${next}${q}`;
			}),
		)
		.replace(/Sheet\.(\d+)(?!\d)/g, (m, id: string) => {
			const next = remap.get(id);
			return next === undefined ? m : `Sheet.${next}`;
		});
}

export interface ShapeInput {
	/** The page key, or the master scope, the shapes belong to (the prefix of their room keys). */
	scope: string;
	/** `shapeKey` -> entry, for one page. */
	shapes: Map<string, ShapeEntry>;
	/** `shapeKey` -> `path` -> entry. */
	cells: Map<string, Map<string, PieceEntry>>;
	/** Persisted renumberings by shape key. */
	renum: ReadonlyMap<string, RenumEntry>;
	/** Masters renumbered in a batch: batch -> old id -> new id. */
	masterRemap: ReadonlyMap<string, ReadonlyMap<string, string>>;
}

/** The shapes of a page in document order (parents first), the XML of the top-level list and the fix-ups. */
export function resolveShapes(input: ShapeInput): {
	shapes: ResolvedShape[];
	xml: string;
	fixups: Fixup[];
} {
	const raw = new Map<string, { entry: ShapeEntry; id: string }>();
	for (const [key, entry] of input.shapes) {
		const id = attrOf(entry.open, 'ID');
		if (id !== undefined) raw.set(key, { entry, id });
	}
	const state = new Map<string, boolean>();
	const valid = (key: string, seen: Set<string> = new Set()): boolean => {
		const known = state.get(key);
		if (known !== undefined) return known;
		const item = raw.get(key);
		const parent = item?.entry.parent;
		let ok = !!item;
		if (ok && parent !== undefined) ok = !seen.has(key) && valid(parent, seen.add(key));
		state.set(key, ok);
		return ok;
	};
	const live = [...raw.keys()].filter((k) => valid(k)).sort(compare);

	const { ids: finalId, fresh: renumbered } = assignIds(
		live.map((key) => {
			const { id } = raw.get(key)!;
			return { key, id, fixed: fixedId(input.renum, key, id) };
		}),
	);
	const batchRemap = new Map<string, Map<string, string>>();
	for (const [key, change] of renumbered) {
		const b = raw.get(key)!.entry.b;
		if (b === undefined) continue;
		const map = batchRemap.get(b) ?? new Map<string, string>();
		map.set(change.from, change.to);
		batchRemap.set(b, map);
	}
	const remapFor = (key: string): Map<string, string> => {
		const out = new Map(batchRemap.get(raw.get(key)!.entry.b ?? '') ?? []);
		const chain: string[] = [];
		for (let k: string | undefined = key; k !== undefined; k = raw.get(k)?.entry.parent)
			chain.unshift(k);
		for (const k of chain) {
			const change = renumbered.get(k);
			if (change) out.set(change.from, change.to);
		}
		return out;
	};

	const children = new Map<string, string[]>();
	for (const key of live) {
		const parent = raw.get(key)!.entry.parent ?? '';
		children.set(parent, [...(children.get(parent) ?? []), key]);
	}
	const sortSiblings = (keys: string[]): string[] =>
		keys.sort((a, b) => raw.get(a)!.entry.pos - raw.get(b)!.entry.pos || compare(a, b));

	const shapes: ResolvedShape[] = [];
	const fixups: Fixup[] = [];
	const render = (key: string): string => {
		const { entry } = raw.get(key)!;
		const id = finalId.get(key)!;
		const remap = remapFor(key);
		let open = setAttr(entry.open, 'ID', id);
		const record: RenumEntry = { ...input.renum.get(key) };
		let dirty = false;
		const change = renumbered.get(key);
		if (change) {
			record.i = [change.from, change.to];
			dirty = true;
		}
		const master = attrOf(open, 'Master');
		const persisted = record.m && record.m[0] === master ? record.m[1] : undefined;
		const batched =
			master !== undefined ? input.masterRemap.get(entry.b ?? '')?.get(master) : undefined;
		const next = persisted ?? batched;
		if (next !== undefined) open = setAttr(open, 'Master', next);
		if (persisted === undefined && batched !== undefined) {
			record.m = [master!, batched];
			dirty = true;
		}
		if (dirty) fixups.push({ map: 'renum', key: `${input.scope}/${key}`, value: record });
		const entries = new Map<string, PieceEntry>();
		for (const [path, piece] of input.cells.get(key) ?? []) {
			const leaf = piece.xml !== undefined && /(^|\/)c\/[^/]*$/.test(path);
			const xml = leaf ? renumberFormulas(piece.xml!, remap) : undefined;
			if (xml !== undefined && xml !== piece.xml) {
				const value = { ...piece, xml };
				entries.set(path, value);
				fixups.push({ map: 'cells', key: `${input.scope}/${key}/${path}`, value });
			} else entries.set(path, piece);
		}
		shapes.push({
			key,
			id,
			...(entry.parent !== undefined ? { parent: entry.parent } : {}),
			open,
			entries,
		});
		const inner = sortSiblings(children.get(key) ?? [])
			.map(render)
			.join('');
		return buildShape(open, entries, inner ? `<Shapes>${inner}</Shapes>` : '');
	};
	const xml = sortSiblings(children.get('') ?? [])
		.map(render)
		.join('');
	return { shapes, xml, fixups };
}
