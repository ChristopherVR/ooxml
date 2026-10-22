// Writes a local drawing into the shared maps, touching only what the person changed. Each entry
// is compared with the content this peer last wrote or applied (its baseline), so a collaborator's
// change that has not been applied here yet is never overwritten by a stale local copy. Items the
// peer knew about and no longer has become tombstones, which win over a concurrent edit of the same
// item. Run it inside one transaction.
import type { DrawingKeys } from './keys';
import { assignPositions } from './order';
import { canon, fingerprint, toBase64 } from './package';
import type { DrawingSnapshot } from './snapshot';
import {
	type PageEntry,
	SCHEMA_VERSION,
	type SharedDrawing,
	bodyEntry,
	pageEntry,
	setIfChanged,
	shapeEntry,
} from './schema';

export const pageId = (key: string): string => `p:${key}`;
export const bodyId = (key: string): string => `b:${key}`;
export const shapeId = (page: string, key: string): string => `s:${page}/${key}`;
export const connectId = (key: string): string => `c:${key}`;
/** Baseline text of a page entry (its position excluded) and of a body. */
export const pageContent = (e: { id: string; back?: string | undefined; xml: string }): string =>
	JSON.stringify([e.id, e.back ?? null, canon(e.xml)]);
export const bodyContent = (e: { shell: string; rels?: string | undefined }): string =>
	JSON.stringify([canon(e.shell), e.rels ?? null]);

/** Writes `snapshot` into `shared`; returns the number of entries changed. */
export function writeSharedDrawing(
	shared: SharedDrawing,
	snapshot: DrawingSnapshot,
	keys: DrawingKeys,
): number {
	let writes = 0;
	const set = (map: SharedDrawing['meta'], key: string, value: unknown): void => {
		if (setIfChanged(map, key, value)) writes++;
	};
	set(shared.meta, 'v', SCHEMA_VERSION);
	set(shared.meta, 'pagesPath', snapshot.pagesPath);
	set(shared.meta, 'pagesOpen', snapshot.pagesOpen);
	set(shared.meta, 'pagesClose', snapshot.pagesClose);

	const now = new Map<string, string>();
	/** Whether the local content differs from what this peer last synced (or from the room). */
	const changed = (id: string, content: string, room: string | undefined): boolean => {
		const base = keys.known.get(id);
		return base === undefined ? room !== content : base !== content;
	};
	const pagePos = assignPositions(
		snapshot.pages.map((p) => p.key),
		(key) => pageEntry(shared.pages.get(key))?.pos,
	);
	for (const page of snapshot.pages) {
		const room = pageEntry(shared.pages.get(page.key));
		const content = pageContent(page);
		now.set(pageId(page.key), content);
		const own = changed(pageId(page.key), content, room && pageContent(room));
		const source: Pick<PageEntry, 'id' | 'back' | 'xml'> = own || !room ? page : room;
		set(shared.pages, page.key, {
			pos: pagePos.get(page.key)!,
			id: source.id,
			...(source.back !== undefined ? { back: source.back } : {}),
			xml: source.xml,
		});
		const roomBody = bodyEntry(shared.bodies.get(page.key));
		const bodyText = bodyContent(page);
		now.set(bodyId(page.key), bodyText);
		const ownBody = changed(bodyId(page.key), bodyText, roomBody && bodyContent(roomBody));
		const body = ownBody || !roomBody ? page : roomBody;
		set(shared.bodies, page.key, {
			shell: body.shell,
			...(body.rels !== undefined ? { rels: body.rels } : {}),
		});
		const shapePos = assignPositions(
			page.shapes.map((s) => s.key),
			(key) => shapeEntry(shared.shapes.get(`${page.key}/${key}`))?.pos,
		);
		for (const shape of page.shapes) {
			const id = shapeId(page.key, shape.key);
			const roomShape = shapeEntry(shared.shapes.get(`${page.key}/${shape.key}`));
			const text = canon(shape.xml);
			now.set(id, text);
			const xml =
				changed(id, text, roomShape && canon(roomShape.xml)) || !roomShape
					? shape.xml
					: roomShape.xml;
			set(shared.shapes, `${page.key}/${shape.key}`, { pos: shapePos.get(shape.key)!, xml });
		}
		for (const connect of page.connects) {
			now.set(connectId(connect.key), '');
			if (!keys.known.has(connectId(connect.key)) || !shared.connects.has(connect.key))
				set(shared.connects, connect.key, connect.entry);
		}
	}

	for (const id of keys.known.keys()) {
		if (now.has(id)) continue;
		const name = id.slice(2);
		if (id.startsWith('p:')) {
			shared.dead.set(name, true);
			for (const map of [shared.pages, shared.bodies]) map.delete(name);
			for (const map of [shared.shapes, shared.connects])
				for (const key of [...map.keys()]) if (key.startsWith(`${name}/`)) map.delete(key);
			keys.forgetPage(name);
			writes++;
		} else if (id.startsWith('s:')) {
			const slash = name.indexOf('/');
			const page = name.slice(0, slash);
			if (!now.has(pageId(page))) continue;
			shared.dead.set(name, true);
			shared.shapes.delete(name);
			keys.forgetShape(page, name.slice(slash + 1));
			writes++;
		} else if (id.startsWith('c:') && shared.connects.has(name)) {
			shared.connects.delete(name);
			writes++;
		}
	}
	keys.known = now;

	for (const [path, data] of snapshot.parts) {
		const print = fingerprint(data);
		if (keys.partPrints.get(path) === print) continue;
		set(shared.parts, path, toBase64(data));
		keys.partPrints.set(path, print);
	}
	for (const path of [...keys.partPrints.keys()])
		if (!snapshot.parts.has(path)) {
			shared.parts.delete(path);
			keys.partPrints.delete(path);
			writes++;
		}
	return writes;
}
