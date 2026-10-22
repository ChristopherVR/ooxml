// Reads the shared maps into a resolved drawing: tombstoned and orphaned items dropped, lists in
// their shared order, and ids that two peers minted concurrently made unique (see read-shapes.ts
// for the shape rules, read-parts.ts for masters and media). Pages collide on `Page/@ID` the same
// way: the lowest key keeps its id and `BackPage` follows. Every peer computes the same result.
import { byPosition } from './order';
import { dirName } from './package';
import type { PieceEntry } from './pieces';
import {
	type ResolvedMaster,
	type ResolvedMedia,
	renameTargets,
	resolveMasters,
	resolveMedia,
} from './read-parts';
import { type ResolvedShape, resolveShapes, setAttr } from './read-shapes';
import {
	type BodyEntry,
	type ConnectEntry,
	type ShapeEntry,
	type SharedDrawing,
	assertSchemaVersion,
	bodyEntry,
	cellEntry,
	connectEntry,
	isSharedDrawingEmpty,
	pageEntry,
	shapeEntry,
	str,
} from './schema';

export { renumberShape } from './read-shapes';
export type { ResolvedMaster, ResolvedMedia, ResolvedShape };

export interface ResolvedConnect {
	key: string;
	entry: ConnectEntry;
	/** Final ids the connect refers to. */
	from: string;
	to: string;
}
export interface ResolvedPage {
	key: string;
	id: string;
	/** Key of the background page. */
	back?: string;
	/** The `Page` element with `Rel` as a marker and `ID`/`BackPage` final. */
	xml: string;
	body: BodyEntry;
	/** Every shape of the page, parents before the shapes nested in them. */
	shapes: ResolvedShape[];
	/** The XML of the top-level shape list. */
	shapesXml: string;
	connects: ResolvedConnect[];
	/** Changes whenever the page part would. */
	signature: string;
}
export interface ResolvedDrawing {
	pagesPath: string;
	pagesOpen: string;
	pagesClose: string;
	pages: ResolvedPage[];
	mastersPath?: string;
	mastersOpen?: string;
	mastersClose?: string;
	masters: ResolvedMaster[];
	media: ResolvedMedia[];
	/** Shared parts as base64, by path. */
	parts: Map<string, string>;
	/** Changes whenever the rebuilt package would. */
	signature: string;
}

const compare = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);
const numeric = (id: string): number => (/^\d+$/.test(id) ? Number(id) : 0);

/** A 64-bit FNV-style hash of the strings, as hex: a change detector, not a security measure. */
export function hashTexts(texts: Iterable<string>): string {
	let a = 0x811c9dc5;
	let b = 0x01000193 ^ 0x9e3779b9;
	for (const text of texts) {
		for (let i = 0; i < text.length; i++) {
			const c = text.charCodeAt(i);
			a = Math.imul(a ^ c, 0x01000193);
			b = Math.imul(b + c, 0x85ebca6b) ^ (b >>> 13);
		}
		a = Math.imul(a ^ 0xff, 0x01000193);
		b = Math.imul(b ^ 0xff, 0xc2b2ae35);
	}
	return `${(a >>> 0).toString(16)}${(b >>> 0).toString(16)}`;
}

/** Splits `key` at its first `depth` slashes; the last part keeps any further slashes. */
const split = (key: string, depth: number): string[] | undefined => {
	const out: string[] = [];
	let from = 0;
	for (let i = 0; i < depth; i++) {
		const slash = key.indexOf('/', from);
		if (slash < 0) return undefined;
		out.push(key.slice(from, slash));
		from = slash + 1;
	}
	out.push(key.slice(from));
	return out;
};

/**
 * The drawing the room holds, or undefined while it is empty. Throws a SchemaVersionError when
 * the room was written with another schema version.
 */
export function readSharedDrawing(shared: SharedDrawing): ResolvedDrawing | undefined {
	assertSchemaVersion(shared);
	if (isSharedDrawingEmpty(shared)) return undefined;
	const pagesPath = str(shared.meta.get('pagesPath'));
	const pagesOpen = str(shared.meta.get('pagesOpen'));
	const pagesClose = str(shared.meta.get('pagesClose'));
	if (pagesPath === undefined || pagesOpen === undefined || pagesClose === undefined)
		return undefined;
	const alive = [...shared.pages]
		.filter(([key]) => !shared.dead.has(key))
		.flatMap(([key, value]) => {
			const entry = pageEntry(value);
			const body = bodyEntry(shared.bodies.get(key));
			return entry && body ? [{ key, pos: entry.pos, entry, body }] : [];
		});
	// Page ids: unique, lowest key keeps its id.
	const pageIds = new Map<string, string>();
	const taken = new Set<string>();
	let maxPage = Math.max(0, ...alive.map((p) => numeric(p.entry.id)));
	for (const page of [...alive].sort((a, b) => compare(a.key, b.key))) {
		let id = page.entry.id;
		if (taken.has(id)) id = String(++maxPage);
		taken.add(id);
		pageIds.set(page.key, id);
	}
	const shapesByPage = new Map<string, Map<string, ShapeEntry>>();
	for (const [full, value] of shared.shapes) {
		const parts = split(full, 1);
		if (!parts || shared.dead.has(full)) continue;
		const entry = shapeEntry(value);
		if (!entry) continue;
		const map = shapesByPage.get(parts[0]!) ?? new Map<string, ShapeEntry>();
		map.set(parts[1]!, entry);
		shapesByPage.set(parts[0]!, map);
	}
	const cellsByPage = new Map<string, Map<string, Map<string, PieceEntry>>>();
	for (const [full, value] of shared.cells) {
		const parts = split(full, 2);
		const entry = parts && cellEntry(value);
		if (!parts || !entry) continue;
		const [page, shape, path] = parts as [string, string, string];
		const shapes = cellsByPage.get(page) ?? new Map<string, Map<string, PieceEntry>>();
		const cells = shapes.get(shape) ?? new Map<string, PieceEntry>();
		cells.set(path, entry);
		shapes.set(shape, cells);
		cellsByPage.set(page, shapes);
	}
	const connectsByPage = new Map<string, { key: string; entry: ConnectEntry }[]>();
	for (const [full, value] of shared.connects) {
		const entry = connectEntry(value);
		const slash = full.indexOf('/');
		if (!entry || slash < 0) continue;
		const list = connectsByPage.get(full.slice(0, slash)) ?? [];
		list.push({ key: full, entry });
		connectsByPage.set(full.slice(0, slash), list);
	}
	const { masters, remap: masterRemap } = resolveMasters(shared);
	const { media, renames } = resolveMedia(shared);
	const mastersPath = str(shared.meta.get('mastersPath'));
	const mastersOpen = str(shared.meta.get('mastersOpen'));
	const mastersClose = str(shared.meta.get('mastersClose'));
	const pagesDir = dirName(pagesPath);

	const pages: ResolvedPage[] = [];
	for (const page of byPosition(alive)) {
		const finalId = pageIds.get(page.key)!;
		const resolved = resolveShapes({
			shapes: shapesByPage.get(page.key) ?? new Map(),
			cells: cellsByPage.get(page.key) ?? new Map(),
			masterRemap,
		});
		const byKey = new Map(resolved.shapes.map((s) => [s.key, s]));
		const connects: ResolvedConnect[] = [];
		for (const item of (connectsByPage.get(page.key) ?? []).sort((a, b) => compare(a.key, b.key))) {
			const { entry } = item;
			const from = entry.fk ? byKey.get(entry.fk) : undefined;
			const to = entry.tk ? byKey.get(entry.tk) : undefined;
			if ((entry.fk && !from) || (entry.tk && !to)) continue;
			connects.push({ key: item.key, entry, from: from?.id ?? entry.fi, to: to?.id ?? entry.ti });
		}
		let xml = page.entry.xml;
		const open = /^<[^>]*>/.exec(xml)?.[0];
		if (open) {
			let next = setAttr(open, 'ID', finalId);
			const back = page.entry.back !== undefined ? pageIds.get(page.entry.back) : undefined;
			if (back !== undefined) next = setAttr(next, 'BackPage', back);
			xml = next + xml.slice(open.length);
		}
		const rename = page.body.b !== undefined ? renames.get(page.body.b) : undefined;
		const body: BodyEntry =
			rename && page.body.rels !== undefined
				? { ...page.body, rels: renameTargets(page.body.rels, pagesDir, rename) }
				: page.body;
		pages.push({
			key: page.key,
			id: finalId,
			...(page.entry.back !== undefined ? { back: page.entry.back } : {}),
			xml,
			body,
			shapes: resolved.shapes,
			shapesXml: resolved.xml,
			connects,
			signature: hashTexts([
				xml,
				body.shell,
				body.rels ?? '\0',
				resolved.xml,
				...connects.flatMap((c) => [c.from, c.to, c.entry.fc, c.entry.tc, c.entry.fp, c.entry.tp]),
			]),
		});
	}
	const mastersDir = mastersPath !== undefined ? dirName(mastersPath) : pagesDir;
	const finalMasters = masters.map((m) => {
		const batch = (shared.masters.get(m.key) as { b?: unknown } | undefined)?.b;
		const rename = typeof batch === 'string' ? renames.get(batch) : undefined;
		return rename && m.rels !== undefined
			? { ...m, rels: renameTargets(m.rels, mastersDir, rename) }
			: m;
	});
	const parts = new Map<string, string>();
	for (const [path, value] of shared.parts) if (typeof value === 'string') parts.set(path, value);
	const signature = hashTexts([
		pagesPath,
		pagesOpen,
		pagesClose,
		mastersPath ?? '',
		mastersOpen ?? '',
		mastersClose ?? '',
		...pages.map((p) => `${p.key}:${p.signature}`),
		...finalMasters.flatMap((m) => [m.key, m.id, m.xml, m.part, m.rels ?? '\0']),
		...media.flatMap((m) => [m.path, m.hash, m.declared.type, String(m.declared.byDefault)]),
		...[...parts].flatMap(([path, data]) => [path, data]),
	]);
	return {
		pagesPath,
		pagesOpen,
		pagesClose,
		pages,
		...(mastersPath !== undefined && mastersOpen !== undefined && mastersClose !== undefined
			? { mastersPath, mastersOpen, mastersClose }
			: {}),
		masters: finalMasters,
		media,
		parts,
		signature,
	};
}
