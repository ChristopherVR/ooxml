// Resolves masters and media from the shared maps.
//
// Masters: one entry each (the shapes inside a master's content are shape and cell entries, see
// read.ts); ids that two peers chose concurrently are made unique like shape ids (read-ids.ts).
// A renumbered master rewrites `Master` on the shapes written in the same batch (see
// read-shapes.ts); shapes of other batches that name the old id are not rewritten, and `Master`
// attributes inside master parts are never rewritten.
//
// Media: one entry per content hash and path. Identical bytes at one path are one entry. Different
// bytes at one path (two peers added different pictures under one name) keep the lowest key at the
// path; the others move to `<name>-<hash prefix><ext>` and the relationship targets of the page
// bodies and masters written in the same batch follow. Identical bytes at different paths stay at
// both paths, so no relationship target of a peer is left dangling.
import { type DeclaredType, relativeTarget, resolveTarget } from './package';
import { masterScope } from './baseline';
import { type Fixup, assignIds, fixedId } from './read-ids';
import { setAttr } from './read-shapes';
import { attrOf } from './scan';
import {
	type MasterEntry,
	type RenumEntry,
	type SharedDrawing,
	masterEntry,
	mediaEntry,
	str,
} from './schema';

export interface MasterBase {
	key: string;
	/** Final id. */
	id: string;
	/** The `Master` element with `Rel` as a marker and `ID` final. */
	xml: string;
	/** Base64 of the master part with its shapes and connects replaced by markers. */
	part: string;
	rels?: string;
}
export interface ResolvedMedia {
	/** Key of its entry in the room. */
	key: string;
	path: string;
	hash: string;
	declared: DeclaredType;
	/** Base64 of the bytes. */
	data: string;
}

const compare = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);

export function resolveMasters(
	shared: SharedDrawing,
	renum: ReadonlyMap<string, RenumEntry>,
): { masters: MasterBase[]; remap: Map<string, Map<string, string>>; fixups: Fixup[] } {
	const alive: { key: string; entry: MasterEntry }[] = [];
	for (const [key, value] of shared.masters) {
		const entry = masterEntry(value);
		if (entry) alive.push({ key, entry });
	}
	const { ids, fresh } = assignIds(
		alive.map(({ key, entry }) => ({
			key,
			id: entry.id,
			fixed: fixedId(renum, masterScope(key), entry.id),
		})),
	);
	const remap = new Map<string, Map<string, string>>();
	const fixups: Fixup[] = [];
	for (const [key, change] of [...fresh].sort((a, b) => compare(a[0], b[0]))) {
		const b = alive.find((m) => m.key === key)!.entry.b;
		if (b !== undefined) {
			const map = remap.get(b) ?? new Map<string, string>();
			map.set(change.from, change.to);
			remap.set(b, map);
		}
		fixups.push({
			map: 'renum',
			key: masterScope(key),
			value: { ...renum.get(masterScope(key)), i: [change.from, change.to] },
		});
	}
	alive.sort((a, b) => a.entry.pos - b.entry.pos || compare(a.key, b.key));
	return {
		masters: alive.map(({ key, entry }) => {
			const id = ids.get(key)!;
			const open = /^<[^>]*>/.exec(entry.xml)?.[0];
			return {
				key,
				id,
				xml: open ? setAttr(open, 'ID', id) + entry.xml.slice(open.length) : entry.xml,
				part: entry.part,
				...(entry.rels !== undefined ? { rels: entry.rels } : {}),
			};
		}),
		remap,
		fixups,
	};
}

export function resolveMedia(shared: SharedDrawing): {
	media: ResolvedMedia[];
	renames: Map<string, Map<string, string>>;
} {
	const items = [...shared.media]
		.flatMap(([key, value]) => {
			const entry = mediaEntry(value);
			const data = entry ? str(shared.blobs.get(entry.hash)) : undefined;
			return entry && data !== undefined ? [{ key, entry, data }] : [];
		})
		.sort((a, b) => compare(a.key, b.key));
	const taken = new Set<string>();
	const first = new Map<string, string>();
	for (const item of items) if (!first.has(item.entry.path)) first.set(item.entry.path, item.key);
	for (const path of first.keys()) taken.add(path);
	const renames = new Map<string, Map<string, string>>();
	const media: ResolvedMedia[] = [];
	for (const { key, entry, data } of items) {
		let path = entry.path;
		if (first.get(entry.path) !== key) {
			const dot = path.lastIndexOf('.');
			const stem = dot > path.lastIndexOf('/') ? path.slice(0, dot) : path;
			const ext = dot > path.lastIndexOf('/') ? path.slice(dot) : '';
			path = `${stem}-${entry.hash.slice(0, 8)}${ext}`;
			for (let n = 2; taken.has(path); n++) path = `${stem}-${entry.hash.slice(0, 8)}-${n}${ext}`;
			taken.add(path);
			if (entry.b !== undefined) {
				const map = renames.get(entry.b) ?? new Map<string, string>();
				map.set(entry.path, path);
				renames.set(entry.b, map);
			}
		}
		media.push({
			key,
			path,
			hash: entry.hash,
			declared: { type: entry.ct, byDefault: entry.d },
			data,
		});
	}
	return { media, renames };
}

/** Rewrites relationship targets that resolve to a renamed part. */
export function renameTargets(
	rels: string,
	baseDir: string,
	renames: ReadonlyMap<string, string>,
): string {
	return rels.replace(/<(?:\w+:)?Relationship\b[^>]*>/g, (tag) => {
		const target = attrOf(tag, 'Target');
		if (target === undefined || attrOf(tag, 'TargetMode') === 'External') return tag;
		const next = renames.get(resolveTarget(baseDir, target));
		if (next === undefined) return tag;
		const value = target.startsWith('/') ? `/${next}` : relativeTarget(baseDir, next);
		return tag.replace(
			/(\sTarget\s*=\s*)(?:"[^"]*"|'[^']*')/,
			(_m, head: string) => `${head}"${value}"`,
		);
	});
}
