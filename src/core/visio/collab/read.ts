// Reads the shared maps into a resolved drawing: tombstoned and orphaned items dropped, lists in
// their shared order, and ids that two peers minted concurrently made unique.
//
// Id rule (deterministic, so every peer renumbers alike): within a page, shapes are visited in key
// order; a shape whose own id or a nested shape's id is already taken keeps the others' ids and
// takes the next numbers above the page's largest id. `Sheet.<n>` references inside the renumbered
// shape and the page's connects follow the new ids; formulas in OTHER shapes that point at a
// renumbered shape are not rewritten (documented limit). Pages collide the same way on `Page/@ID`.
import { byPosition } from './order';
import { shapeIds } from './snapshot';
import {
	type BodyEntry,
	type ConnectEntry,
	type SharedDrawing,
	bodyEntry,
	connectEntry,
	isSharedDrawingEmpty,
	pageEntry,
	shapeEntry,
	str,
} from './schema';

export interface ResolvedShape {
	key: string;
	/** Id of the outermost shape, after renumbering. */
	id: string;
	xml: string;
	/** Original id -> final id, only for renumbered shapes. */
	remap: Map<string, string>;
	/** Final ids of this shape and everything nested in it. */
	ids: string[];
}
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
	shapes: ResolvedShape[];
	connects: ResolvedConnect[];
}
export interface ResolvedDrawing {
	pagesPath: string;
	pagesOpen: string;
	pagesClose: string;
	pages: ResolvedPage[];
	/** Shared parts as base64, by path. */
	parts: Map<string, string>;
}

const compare = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);
const numeric = (id: string): number => (/^\d+$/.test(id) ? Number(id) : 0);

function setAttr(open: string, name: string, value: string): string {
	return open.replace(
		new RegExp(String.raw`(\s${name}\s*=\s*)(?:"[^"]*"|'[^']*')`),
		(_m, head: string) => `${head}"${value}"`,
	);
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

export function readSharedDrawing(shared: SharedDrawing): ResolvedDrawing | undefined {
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
	const shapesByPage = new Map<string, { key: string; pos: number; xml: string }[]>();
	for (const [full, value] of shared.shapes) {
		const slash = full.indexOf('/');
		if (slash < 0 || shared.dead.has(full)) continue;
		const entry = shapeEntry(value);
		if (!entry) continue;
		const page = full.slice(0, slash);
		const list = shapesByPage.get(page) ?? [];
		list.push({ key: full.slice(slash + 1), pos: entry.pos, xml: entry.xml });
		shapesByPage.set(page, list);
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
	const pages: ResolvedPage[] = [];
	for (const page of byPosition(alive)) {
		const finalId = pageIds.get(page.key)!;
		const shapes = resolveShapes(shapesByPage.get(page.key) ?? []);
		const byKey = new Map(shapes.map((s) => [s.key, s]));
		const connects: ResolvedConnect[] = [];
		for (const item of (connectsByPage.get(page.key) ?? []).sort((a, b) => compare(a.key, b.key))) {
			const { entry } = item;
			const from = entry.fk ? byKey.get(entry.fk) : undefined;
			const to = entry.tk ? byKey.get(entry.tk) : undefined;
			if ((entry.fk && !from) || (entry.tk && !to)) continue;
			connects.push({
				key: item.key,
				entry,
				from: from?.remap.get(entry.fi) ?? entry.fi,
				to: to?.remap.get(entry.ti) ?? entry.ti,
			});
		}
		let xml = page.entry.xml;
		const open = /^<[^>]*>/.exec(xml)?.[0];
		if (open) {
			let next = setAttr(open, 'ID', finalId);
			const back = page.entry.back !== undefined ? pageIds.get(page.entry.back) : undefined;
			if (back !== undefined) next = setAttr(next, 'BackPage', back);
			xml = next + xml.slice(open.length);
		}
		pages.push({
			key: page.key,
			id: finalId,
			...(page.entry.back !== undefined ? { back: page.entry.back } : {}),
			xml,
			body: page.body,
			shapes,
			connects,
		});
	}
	const parts = new Map<string, string>();
	for (const [path, value] of shared.parts) if (typeof value === 'string') parts.set(path, value);
	return { pagesPath, pagesOpen, pagesClose, pages, parts };
}

function resolveShapes(entries: { key: string; pos: number; xml: string }[]): ResolvedShape[] {
	const parsed = entries.map((entry) => ({ ...entry, ids: shapeIds(entry.xml) }));
	let max = Math.max(0, ...parsed.flatMap((p) => p.ids.map(numeric)));
	const used = new Set<string>();
	const done = new Map<string, ResolvedShape>();
	for (const item of [...parsed].sort((a, b) => compare(a.key, b.key))) {
		const remap = new Map<string, string>();
		for (const id of item.ids) {
			if (used.has(id) || remap.has(id)) remap.set(id, String(++max));
		}
		const ids = item.ids.map((id) => remap.get(id) ?? id);
		for (const id of ids) used.add(id);
		done.set(item.key, {
			key: item.key,
			id: ids[0] ?? '',
			xml: renumberShape(item.xml, remap),
			remap,
			ids,
		});
	}
	return byPosition(parsed).map((item) => done.get(item.key)!);
}
