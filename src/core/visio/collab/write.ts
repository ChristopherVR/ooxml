// Writes a local drawing into the shared maps, touching only what the person changed. Each entry
// is compared with the content this peer last wrote or applied (its baseline), so a collaborator's
// change that has not been applied here yet is never overwritten by a stale local copy. Items the
// peer knew about and no longer has are removed: pages and shapes become tombstones, which win
// over a concurrent edit of the same item; cells, masters and media are simply deleted, so a
// concurrent edit of one of those survives. Run it inside one transaction.
import type * as Y from 'yjs';
import {
	bodyContent,
	bodyId,
	connectId,
	masterId,
	masterScope,
	pageContent,
	pageId,
} from './baseline';
import type { DrawingKeys } from './keys';
import { assignPositions } from './order';
import {
	type BodyEntry,
	type PageEntry,
	SCHEMA_VERSION,
	type SharedDrawing,
	assertSchemaVersion,
	bodyEntry,
	pageEntry,
	setIfChanged,
} from './schema';
import type { DrawingSnapshot } from './snapshot';
import { syncMasters, syncMedia, syncParts } from './write-parts';
import { type WriteContext, syncConnects, syncEntry, syncShapes } from './write-shapes';

export { bodyId, connectId, pageContent, bodyContent, pageId };

/** Removes the shapes, cells, connects and renumberings of a page or master scope. */
function removeScope(shared: SharedDrawing, name: string): void {
	shared.renum.delete(name);
	for (const map of [shared.shapes, shared.cells, shared.connects, shared.renum])
		for (const key of [...map.keys()]) if (key.startsWith(`${name}/`)) map.delete(key);
}

function removePage(shared: SharedDrawing, name: string): void {
	shared.dead.set(name, true);
	for (const map of [shared.pages, shared.bodies]) map.delete(name);
	removeScope(shared, name);
}

/** Writes `snapshot` into `shared`; returns the number of entries changed. Throws on another schema version. */
export function writeSharedDrawing(
	shared: SharedDrawing,
	snapshot: DrawingSnapshot,
	keys: DrawingKeys,
): number {
	assertSchemaVersion(shared);
	let writes = 0;
	const write = (map: Y.Map<unknown>, key: string, value: unknown): void => {
		if (setIfChanged(map, key, value)) writes++;
	};
	const now = new Map<string, string>();
	const ctx: WriteContext = { shared, keys, now, batch: keys.batch(), write };
	write(shared.meta, 'v', SCHEMA_VERSION);
	write(shared.meta, 'pagesPath', snapshot.pagesPath);
	write(shared.meta, 'pagesOpen', snapshot.pagesOpen);
	write(shared.meta, 'pagesClose', snapshot.pagesClose);
	if (snapshot.mastersPath !== undefined) {
		write(shared.meta, 'mastersPath', snapshot.mastersPath);
		write(shared.meta, 'mastersOpen', snapshot.mastersOpen);
		write(shared.meta, 'mastersClose', snapshot.mastersClose);
	}

	const pagePos = assignPositions(
		snapshot.pages.map((p) => p.key),
		(key) => pageEntry(shared.pages.get(key))?.pos,
	);
	for (const page of snapshot.pages) {
		const room = pageEntry(shared.pages.get(page.key));
		syncEntry<PageEntry>(
			ctx,
			shared.pages,
			page.key,
			pageId(page.key),
			pageContent(page),
			room,
			() => pageContent(room!),
			() => ({
				pos: pagePos.get(page.key)!,
				id: page.id,
				...(page.back !== undefined ? { back: page.back } : {}),
				xml: page.xml,
			}),
		);
		const roomBody = bodyEntry(shared.bodies.get(page.key));
		syncEntry<BodyEntry>(
			ctx,
			shared.bodies,
			page.key,
			bodyId(page.key),
			bodyContent(page),
			roomBody,
			() => bodyContent(roomBody!),
			() => ({ shell: page.shell, ...(page.rels !== undefined ? { rels: page.rels } : {}) }),
			true,
		);
		syncShapes(ctx, page);
		syncConnects(ctx, page.connects);
	}
	syncMasters(ctx, snapshot);
	syncMedia(ctx, snapshot);

	for (const id of keys.known.keys()) {
		if (now.has(id)) continue;
		const name = id.slice(2);
		switch (id.slice(0, 2)) {
			case 'p:':
				removePage(shared, name);
				keys.forgetPage(name);
				writes++;
				break;
			case 's:': {
				const slash = name.indexOf('/');
				const page = name.slice(0, slash);
				if (!now.has(page.startsWith(masterScope('')) ? masterId(page.slice(2)) : pageId(page)))
					break;
				shared.dead.set(name, true);
				shared.shapes.delete(name);
				keys.forgetShape(page, name.slice(slash + 1));
				writes++;
				break;
			}
			case 'f:':
				if (shared.cells.has(name)) {
					shared.cells.delete(name);
					writes++;
				}
				break;
			case 'c:':
				if (shared.connects.has(name)) {
					shared.connects.delete(name);
					writes++;
				}
				break;
			case 'm:':
				if (shared.masters.has(name)) {
					shared.masters.delete(name);
					removeScope(shared, masterScope(name));
					writes++;
				}
				keys.forgetMaster(name);
				break;
			case 'i:':
				if (shared.media.has(name)) {
					shared.media.delete(name);
					writes++;
				}
				break;
		}
	}
	keys.known = now;
	writes += syncParts(ctx, snapshot);
	return writes;
}
