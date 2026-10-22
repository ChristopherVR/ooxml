// Resolves the shapes of one page from the shared maps: tombstoned and orphaned shapes dropped,
// siblings in their shared order, ids that two peers minted concurrently made unique, and the XML
// of each top-level shape rebuilt from its shell and cell entries in a fixed order.
//
// Id rule (deterministic, so every peer renumbers alike): shapes are visited in key order; a shape
// whose id is already taken takes the next number above the page's largest id and keeps its key.
// `Sheet.<old>` in the F attribute of a cell is rewritten to the new id in (a) the renumbered
// shape and every shape nested in it, and (b) every shape written in the same batch as the
// renumbered one (`b`), whose author numbered them together. It is not rewritten in other
// shapes: a reference written after its author had adopted the new id cannot be told from a
// reference to the shape that kept the old id. Masters renumbered the same way rewrite `Master`
// on the shapes of their batch.
import { buildShape, type PieceEntry } from './pieces';
import { attrOf } from './scan';
import type { ShapeEntry } from './schema';

export interface ResolvedShape {
	key: string;
	/** Final id, after renumbering. */
	id: string;
	parent?: string;
	open: string;
	entries: Map<string, PieceEntry>;
}

const compare = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);
const numeric = (id: string): number => (/^\d+$/.test(id) ? Number(id) : 0);

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
	/** `shapeKey` -> entry, for one page. */
	shapes: Map<string, ShapeEntry>;
	/** `shapeKey` -> `path` -> entry. */
	cells: Map<string, Map<string, PieceEntry>>;
	/** Masters renumbered in a batch: batch -> old id -> new id. */
	masterRemap: ReadonlyMap<string, ReadonlyMap<string, string>>;
}

/** The shapes of a page in document order (parents first) and the XML of the top-level list. */
export function resolveShapes(input: ShapeInput): { shapes: ResolvedShape[]; xml: string } {
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

	let max = Math.max(0, ...live.map((k) => numeric(raw.get(k)!.id)));
	const used = new Set<string>();
	const finalId = new Map<string, string>();
	const renumbered = new Map<string, { from: string; to: string }>();
	for (const key of live) {
		const { id } = raw.get(key)!;
		let next = id;
		if (used.has(id)) {
			next = String(++max);
			renumbered.set(key, { from: id, to: next });
		}
		used.add(next);
		finalId.set(key, next);
	}
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
	const render = (key: string): string => {
		const { entry } = raw.get(key)!;
		const id = finalId.get(key)!;
		const remap = remapFor(key);
		let open = setAttr(entry.open, 'ID', id);
		const master = attrOf(open, 'Master');
		const next =
			master !== undefined ? input.masterRemap.get(entry.b ?? '')?.get(master) : undefined;
		if (next !== undefined) open = setAttr(open, 'Master', next);
		const entries = new Map<string, PieceEntry>();
		for (const [path, piece] of input.cells.get(key) ?? []) {
			const leaf = piece.xml !== undefined && /(^|\/)c\/[^/]*$/.test(path);
			entries.set(path, leaf ? { ...piece, xml: renumberFormulas(piece.xml!, remap) } : piece);
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
	return { shapes, xml };
}
