// Reads the shared maps into a resolved drawing: tombstoned and orphaned items dropped, lists in
// their shared order, and ids that two peers minted concurrently made unique (see read-shapes.ts
// for the shape rules, read-parts.ts for masters and media). Pages collide on `Page/@ID` the same
// way: the lowest key keeps its id and `BackPage` follows. Every peer computes the same result.
import { masterScope } from './baseline';
import { byPosition } from './order';
import { dirName } from './package';
import type { PieceEntry } from './pieces';
import { type Fixup, assignIds, compare, fixedId, readRenum } from './read-ids';
import {
	type ResolvedConnect,
	connectSignature,
	hashTexts,
	resolveConnects,
	split,
} from './read-util';
import {
	type MasterBase,
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
	type RenumEntry,
	shapeEntry,
	str,
} from './schema';

export { renumberShape } from './read-shapes';
export type { ResolvedMedia, ResolvedShape };
export type { Fixup, ResolvedConnect };

export interface ResolvedMaster extends MasterBase {
	/** The shapes inside the master's content, parents before the shapes nested in them. */
	shapes: ResolvedShape[];
	/** The XML of the top-level shape list. */
	shapesXml: string;
	connects: ResolvedConnect[];
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
	/** Changes whenever the page's shapes, connects or part shell do (not its entry in the page list). */
	contentSignature: string;
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
	/**
	 * Changes whenever anything but the content of the pages does: the page list (keys, ids, order,
	 * entries), masters, media and the shared parts. While it stays the same, only the pages whose
	 * `contentSignature` changed differ.
	 */
	baseSignature: string;
	/** Writes that persist the renumberings above, see fixups.ts. Empty once ids are stable. */
	fixups: Fixup[];
}

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
	const renum = readRenum(shared);
	// Page ids: unique, a persisted renumbering first, then the lowest key keeps its id.
	const fixups: Fixup[] = [];
	const { ids: pageIds, fresh: freshPages } = assignIds(
		alive.map((p) => ({
			key: p.key,
			id: p.entry.id,
			fixed: fixedId(renum, p.key, p.entry.id),
		})),
	);
	for (const [key, change] of [...freshPages].sort((a, b) => compare(a[0], b[0])))
		fixups.push({
			map: 'renum',
			key,
			value: { ...renum.get(key), i: [change.from, change.to] },
		});
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
	const renumByScope = new Map<string, Map<string, RenumEntry>>();
	for (const [full, entry] of renum) {
		const parts = split(full, 1);
		if (!parts) continue;
		const map = renumByScope.get(parts[0]!) ?? new Map<string, RenumEntry>();
		map.set(parts[1]!, entry);
		renumByScope.set(parts[0]!, map);
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
	const { masters, remap: masterRemap, fixups: masterFixups } = resolveMasters(shared, renum);
	fixups.push(...masterFixups);
	const { media, renames } = resolveMedia(shared);
	const mastersPath = str(shared.meta.get('mastersPath'));
	const mastersOpen = str(shared.meta.get('mastersOpen'));
	const mastersClose = str(shared.meta.get('mastersClose'));
	const pagesDir = dirName(pagesPath);

	const pages: ResolvedPage[] = [];
	for (const page of byPosition(alive)) {
		const finalId = pageIds.get(page.key)!;
		const resolved = resolveShapes({
			scope: page.key,
			shapes: shapesByPage.get(page.key) ?? new Map(),
			cells: cellsByPage.get(page.key) ?? new Map(),
			renum: renumByScope.get(page.key) ?? new Map(),
			masterRemap,
		});
		fixups.push(...resolved.fixups);
		const connects = resolveConnects(connectsByPage.get(page.key), resolved.shapes);
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
		const contentSignature = hashTexts([
			body.shell,
			body.rels ?? '\0',
			resolved.xml,
			...connectSignature(connects),
		]);
		pages.push({
			key: page.key,
			id: finalId,
			...(page.entry.back !== undefined ? { back: page.entry.back } : {}),
			xml,
			body,
			shapes: resolved.shapes,
			shapesXml: resolved.xml,
			connects,
			signature: hashTexts([xml, contentSignature]),
			contentSignature,
		});
	}
	const mastersDir = mastersPath !== undefined ? dirName(mastersPath) : pagesDir;
	const finalMasters: ResolvedMaster[] = masters.map((m) => {
		const batch = (shared.masters.get(m.key) as { b?: unknown } | undefined)?.b;
		const rename = typeof batch === 'string' ? renames.get(batch) : undefined;
		const scope = masterScope(m.key);
		const resolved = resolveShapes({
			scope,
			shapes: shapesByPage.get(scope) ?? new Map(),
			cells: cellsByPage.get(scope) ?? new Map(),
			renum: renumByScope.get(scope) ?? new Map(),
			masterRemap: new Map(),
		});
		fixups.push(...resolved.fixups);
		return {
			...m,
			...(rename && m.rels !== undefined
				? { rels: renameTargets(m.rels, mastersDir, rename) }
				: {}),
			shapes: resolved.shapes,
			shapesXml: resolved.xml,
			connects: resolveConnects(connectsByPage.get(scope), resolved.shapes),
		};
	});
	const parts = new Map<string, string>();
	for (const [path, value] of shared.parts) if (typeof value === 'string') parts.set(path, value);
	const baseSignature = hashTexts([
		pagesPath,
		pagesOpen,
		pagesClose,
		mastersPath ?? '',
		mastersOpen ?? '',
		mastersClose ?? '',
		...pages.map((p) => `${p.key}:${p.id}:${p.xml}`),
		...finalMasters.flatMap((m) => [
			m.key,
			m.id,
			m.xml,
			m.part,
			m.rels ?? '\0',
			m.shapesXml,
			...connectSignature(m.connects),
		]),
		...media.flatMap((m) => [m.path, m.hash, m.declared.type, String(m.declared.byDefault)]),
		...[...parts].flatMap(([path, data]) => [path, data]),
	]);
	const signature = hashTexts([
		baseSignature,
		...pages.map((p) => `${p.key}:${p.contentSignature}`),
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
		baseSignature,
		fixups,
	};
}
