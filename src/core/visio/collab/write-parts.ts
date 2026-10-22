// Writes the non-page parts: masters (one entry each), media (one entry per content hash and path,
// the bytes stored once per hash) and every other part whole.
import { masterContent, masterId, masterScope, mediaId } from './baseline';
import { assignPositions } from './order';
import { fingerprint, fromBase64, toBase64 } from './package';
import { type MasterEntry, masterEntry, mediaEntry } from './schema';
import type { DrawingSnapshot } from './snapshot';
import { type WriteContext, syncConnects, syncEntry, syncShapes } from './write-shapes';

export const mediaKey = (hash: string, path: string): string => `${hash}|${path}`;

export function syncMasters(ctx: WriteContext, snapshot: DrawingSnapshot): void {
	const { shared } = ctx;
	const pos = assignPositions(
		snapshot.masters.map((m) => m.key),
		(key) => masterEntry(shared.masters.get(key))?.pos,
	);
	for (const master of snapshot.masters) {
		const room = masterEntry(shared.masters.get(master.key));
		syncEntry<MasterEntry>(
			ctx,
			shared.masters,
			master.key,
			masterId(master.key),
			masterContent({ ...master, print: fingerprint(master.part) }),
			room,
			() =>
				masterContent({
					...room!,
					print: fingerprint(fromBase64(room!.part)),
				}),
			() => ({
				pos: pos.get(master.key)!,
				id: master.id,
				xml: master.xml,
				part: toBase64(master.part),
				...(master.rels !== undefined ? { rels: master.rels } : {}),
			}),
			true,
		);
		syncShapes(ctx, { key: masterScope(master.key), shapes: master.shapes });
		syncConnects(ctx, master.connects);
	}
}

export function syncMedia(ctx: WriteContext, snapshot: DrawingSnapshot): void {
	const { shared } = ctx;
	for (const item of snapshot.media) {
		const known = ctx.keys.mediaKeys.get(item.path);
		const key = known?.startsWith(`${item.hash}|`) ? known : mediaKey(item.hash, item.path);
		ctx.now.set(mediaId(key), '');
		if (!mediaEntry(shared.media.get(key))) {
			ctx.write(shared.media, key, {
				path: item.path,
				hash: item.hash,
				ct: item.ct,
				d: item.d,
				b: ctx.batch,
			});
			shared.blobs.set(item.hash, toBase64(item.data));
		} else if (!shared.blobs.has(item.hash)) shared.blobs.set(item.hash, toBase64(item.data));
	}
}

/** Writes every other part whole, last write wins; returns the number of deletions. */
export function syncParts(ctx: WriteContext, snapshot: DrawingSnapshot): number {
	const { shared, keys } = ctx;
	for (const [path, data] of snapshot.parts) {
		const print = fingerprint(data);
		if (keys.partPrints.get(path) === print) continue;
		ctx.write(shared.parts, path, toBase64(data));
		keys.partPrints.set(path, print);
	}
	let removed = 0;
	for (const path of [...keys.partPrints.keys()])
		if (!snapshot.parts.has(path)) {
			shared.parts.delete(path);
			keys.partPrints.delete(path);
			removed++;
		}
	return removed;
}
