import * as Y from 'yjs';
import { mergeSnapshotMaps, mergeSnapshotValue, snapshotEqual } from '../../collab/merge-values';
import { composeDrawing } from './compose';
import { DrawingKeys } from './keys';
import { fromBase64, isLocalOnlyPart, toBase64 } from './package';
import { readSharedDrawing } from './read';
import { sharedDrawing } from './schema';
import { type DrawingSnapshot, readZip, snapshotDrawing } from './snapshot';
import { writeSharedDrawing } from './write';

type MergeResult =
	| { status: 'merged'; bytes: Uint8Array }
	| { status: 'unresolved'; issues: readonly string[] };

/** Detect deletion against changes under an otherwise unchanged parent in the flat Visio maps. */
function deletionIssues(baseDoc: Y.Doc, localDoc: Y.Doc, remoteDoc: Y.Doc): string[] {
	const issues: string[] = [];
	const baseline = sharedDrawing(baseDoc);
	const content = (doc: Y.Doc, prefix: string, shape = false) => {
		const shared = sharedDrawing(doc);
		const selected = new Set([prefix]);
		if (shape)
			for (let count = -1; count !== selected.size;) {
				count = selected.size;
				const page = prefix.slice(0, prefix.indexOf('/'));
				for (const [key, raw] of shared.shapes) {
					const parent = (raw as { parent?: string }).parent;
					if (parent && selected.has(`${page}/${parent}`)) selected.add(key);
				}
			}
		const result: Record<string, unknown> = {};
		for (const [name, map] of Object.entries(shared))
			for (const [key, value] of map)
				if (
					[...selected].some((item) => key === item || key.startsWith(`${item}/`)) ||
					(shape &&
						key.startsWith(`${prefix.slice(0, prefix.indexOf('/'))}/`) &&
						[...selected].some((item) => {
							const id = item.slice(item.indexOf('/') + 1);
							const entry = value as { fk?: string; tk?: string; r?: string[] };
							return (
								(name === 'connects' && (entry.fk === id || entry.tk === id)) ||
								(name === 'cells' && entry.r?.includes(id))
							);
						}))
				)
					result[`${name}/${key}`] = value;
		return result;
	};
	for (const [leftDoc, rightDoc] of [
		[localDoc, remoteDoc],
		[remoteDoc, localDoc],
	] as const) {
		const left = sharedDrawing(leftDoc),
			right = sharedDrawing(rightDoc);
		for (const [name, baseMap, leftMap, rightMap] of [
			['page', baseline.pages, left.pages, right.pages],
			['master', baseline.masters, left.masters, right.masters],
			['shape', baseline.shapes, left.shapes, right.shapes],
		] as const)
			for (const key of baseMap.keys())
				if (
					!leftMap.has(key) &&
					rightMap.has(key) &&
					!snapshotEqual(
						content(baseDoc, name === 'master' ? `m~${key}` : key, name === 'shape'),
						content(rightDoc, name === 'master' ? `m~${key}` : key, name === 'shape'),
					)
				)
					issues.push(`A deleted Visio ${name} has concurrent changes: ${key}.`);
	}
	return issues;
}

/**
 * Combines independent saved VSDX edits through the existing page/shape/cell Yjs mapping.
 * Saved snapshots contain no operation history: overlapping XML leaves remain unresolved.
 */
export async function mergeDrawingSnapshots(
	base: Uint8Array,
	local: Uint8Array,
	remote: Uint8Array,
): Promise<MergeResult> {
	const docs: Y.Doc[] = [];
	try {
		const seed = new Y.Doc();
		docs.push(seed);
		const seedKeys = new DrawingKeys('snapshot-base');
		seedKeys.seeding = true;
		const baseline = await snapshotDrawing(base, seedKeys);
		seed.transact(() => writeSharedDrawing(sharedDrawing(seed), baseline, seedKeys));
		const resolved = readSharedDrawing(sharedDrawing(seed));
		if (!resolved)
			return { status: 'unresolved', issues: ['The Visio baseline could not be mapped.'] };
		const branches: { doc: Y.Doc; snapshot: DrawingSnapshot }[] = [];
		for (const [index, bytes] of [local, remote].entries()) {
			const doc = new Y.Doc();
			docs.push(doc);
			Y.applyUpdate(doc, Y.encodeStateAsUpdate(seed));
			const keys = new DrawingKeys(`snapshot-${index}`);
			(await composeDrawing(resolved, base, keys)).commit();
			const snapshot = await snapshotDrawing(bytes, keys);
			doc.transact(() => writeSharedDrawing(sharedDrawing(doc), snapshot, keys));
			branches.push({ doc, snapshot });
		}
		const left = branches[0]!,
			right = branches[1]!;
		const issues = deletionIssues(seed, left.doc, right.doc);
		for (const snapshot of [baseline, left.snapshot, right.snapshot])
			if ([...snapshot.parts.keys()].some((path) => path.startsWith('_xmlsignatures/')))
				issues.push('Signed drawings require explicit signature handling before merging.');
		mergeSnapshotValue(
			...([baseline, left.snapshot, right.snapshot].map((snapshot) =>
				Object.fromEntries(snapshot.media.map((media) => [media.path, media.hash])),
			) as [unknown, unknown, unknown]),
			'Visio media parts',
			issues,
		);
		const merged = mergeSnapshotMaps(seed, left.doc, right.doc, [...seed.share.keys()]);
		docs.push(merged.doc);
		issues.push(...merged.issues);
		const drawing = readSharedDrawing(sharedDrawing(merged.doc));
		if (!drawing) issues.push('The merged Visio drawing could not be mapped.');
		// Window settings and thumbnails intentionally do not participate in live collaboration.
		// A saved-file merge must preserve their independent edits rather than dropping them.
		const zips = await Promise.all([base, local, remote].map((bytes) => readZip(bytes)));
		const carrier = zips[2]!;
		for (const path of new Set(
			zips.flatMap((zip) => Object.keys(zip.files)).filter(isLocalOnlyPart),
		)) {
			const values = await Promise.all(
				zips.map(async (zip) => {
					const entry = zip.file(path);
					return entry ? toBase64(await entry.async('uint8array')) : undefined;
				}),
			);
			const value = mergeSnapshotValue(values[0], values[1], values[2], path, issues);
			if (value === undefined) carrier.remove(path);
			else carrier.file(path, fromBase64(value as string));
		}
		if (issues.length || !drawing) return { status: 'unresolved', issues };
		const bytes = await carrier.generateAsync({ type: 'uint8array' });
		return {
			status: 'merged',
			bytes: (await composeDrawing(drawing, bytes, new DrawingKeys('snapshot-result'))).bytes,
		};
	} catch {
		return {
			status: 'unresolved',
			issues: ['The VSDX snapshots could not be safely parsed or combined.'],
		};
	} finally {
		for (const doc of docs) doc.destroy();
	}
}
