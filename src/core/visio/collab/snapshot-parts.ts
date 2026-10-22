// Snapshot of the masters, the media and every other part of a local package.
import type JSZip from 'jszip';
import { masterScope } from './baseline';
import type { DrawingKeys } from './keys';
import {
	CONTENT_TYPES,
	MASTER_CONTENT_TYPE,
	PAGE_CONTENT_TYPE,
	contentHash,
	declaredType,
	decoder,
	encoder,
	isLocalOnlyPart,
	isMediaPart,
	relsPathOf,
	stripOverrides,
} from './package';
import { splitContents } from './snapshot-contents';
import type { Layout } from './snapshot-io';
import { readLocalItems, withMark, zipText } from './snapshot-io';
import type { MasterSnapshot, MediaSnapshot } from './snapshot';

export async function snapshotMasters(
	zip: JSZip,
	layout: Layout,
	keys: DrawingKeys,
	derived: Set<string>,
) {
	const masters: MasterSnapshot[] = [];
	let mastersOpen: string | undefined;
	let mastersClose: string | undefined;
	if (layout.mastersPath) {
		const listXml = (await zipText(zip, layout.mastersPath)) ?? '';
		const list = await readLocalItems(zip, layout.mastersPath, listXml, 'Masters', 'Master');
		if (list) {
			derived.add(layout.mastersPath).add(relsPathOf(layout.mastersPath));
			mastersOpen = listXml.slice(0, list.box.openEnd);
			mastersClose = listXml.slice(list.box.closeStart);
			for (const item of list.items) {
				const content = await zipText(zip, item.path);
				if (content === undefined) continue;
				derived.add(item.path).add(relsPathOf(item.path));
				const key = keys.masterKey(item.id);
				keys.masterPaths.set(key, item.path);
				const rels = await zipText(zip, relsPathOf(item.path));
				const { shell, shapes, connects } = splitContents(content, masterScope(key), keys);
				masters.push({
					key,
					id: item.id,
					xml: withMark(item),
					path: item.path,
					part: encoder.encode(shell),
					...(rels !== undefined ? { rels } : {}),
					shapes,
					connects,
				});
			}
		}
	}

	return { masters, mastersOpen, mastersClose };
}

export async function snapshotRest(zip: JSZip, derived: ReadonlySet<string>, limit: number) {
	const parts = new Map<string, Uint8Array>();
	const media: MediaSnapshot[] = [];
	const typesXml = (await zipText(zip, CONTENT_TYPES)) ?? '';
	const mediaPaths = new Set<string>();
	let total = 0;
	for (const path of Object.keys(zip.files)) {
		const file = zip.files[path]!;
		if (file.dir || derived.has(path) || isLocalOnlyPart(path)) continue;
		const data = await file.async('uint8array');
		total += data.length;
		if (total > limit) throw new Error('The package parts exceed the shared size limit.');
		const declared = isMediaPart(path) ? declaredType(typesXml, path) : undefined;
		if (declared) {
			mediaPaths.add(path);
			media.push({
				path,
				hash: contentHash(data),
				ct: declared.type,
				d: declared.byDefault,
				data,
			});
		} else parts.set(path, data);
	}
	const types = parts.get(CONTENT_TYPES);
	if (types)
		parts.set(
			CONTENT_TYPES,
			encoder.encode(
				stripOverrides(decoder.decode(types), [PAGE_CONTENT_TYPE, MASTER_CONTENT_TYPE], mediaPaths),
			),
		);
	return { parts, media };
}
