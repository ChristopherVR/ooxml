// Resolves the shapes of one page from the shared maps: tombstoned and orphaned shapes dropped,
// siblings in their shared order, ids that two peers minted concurrently made unique, and the XML
// of each top-level shape rebuilt from its shell and cell entries in a fixed order.
//
// Id rule (deterministic, so every peer renumbers alike, see read-ids.ts): a shape whose renumbering
// is persisted keeps it; the others are visited in key order, and one whose id is already taken
// takes the next number above the page's largest id and keeps its key. A peer that can write then
// persists the result (fixups.ts), so the id no longer depends on what is added later.
// `Sheet.<n>` in the F attribute of a cell is resolved by shape key, not by number: the entry
// records which shape each referenced id meant to its author (`r`, see pieces.ts), and the reader
// writes the final id of that shape. The room is never rewritten, so the author's entry (and their
// undo of it) is untouched, and the result is a function of the room alone. Masters renumbered in
// a batch rewrite `Master` on the shapes of their batch.
import { buildShape, type PieceEntry } from './pieces';
import { type Fixup, assignIds, compare, fixedId, renumRecord } from './read-ids';
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
/** The final id of each shape a cell entry names in its `r` (id, key pairs); keys that are gone are skipped. */
function referenced(
	r: readonly string[],
	finalId: ReadonlyMap<string, string>,
): Map<string, string> {
	const out = new Map<string, string>();
	for (let i = 0; i + 1 < r.length; i += 2) {
		const next = finalId.get(r[i + 1]!);
		if (next !== undefined) out.set(r[i]!, next);
	}
	return out;
}

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

	const {
		ids: finalId,
		fresh: renumbered,
		claims,
	} = assignIds(
		live.map((key) => {
			const { id } = raw.get(key)!;
			return { key, id, fixed: fixedId(input.renum, key, id) };
		}),
	);
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
		let open = setAttr(entry.open, 'ID', id);
		const idRecord = renumRecord(
			input.renum,
			`${input.scope}/${key}`,
			{ id: raw.get(key)!.id },
			renumbered.get(key),
			claims.has(key),
		);
		const record: RenumEntry = { ...(idRecord ?? input.renum.get(key)) };
		let dirty = idRecord !== undefined;
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
			const leaf = piece.xml !== undefined && piece.r !== undefined;
			const xml = leaf ? renumberFormulas(piece.xml!, referenced(piece.r!, finalId)) : undefined;
			entries.set(path, xml !== undefined && xml !== piece.xml ? { ...piece, xml } : piece);
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
