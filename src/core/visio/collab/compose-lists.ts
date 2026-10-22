// Places the items of a list part (pages, masters) into a package: each item keeps the part name
// it has locally or gets the next free one, and the list part, its relationships and the item
// parts are written. Relationships of other types in the local list rels are kept.
import type JSZip from 'jszip';
import { REL_NS, dirName, encoder, parseRelationships, relsPathOf } from './package';
import { escapeAttr } from './scan';
import { REL_MARK } from './schema';
import { zipText } from './snapshot';

export interface ListItem {
	key: string;
	/** The list element with {@link REL_MARK} where its `Rel` goes. */
	xml: string;
	part: Uint8Array;
	/** Relationships of the item part, verbatim. */
	rels?: string;
}

export interface ListOptions {
	listPath: string;
	open: string;
	close: string;
	relType: string;
	/** File name stem of new parts: `page` gives `page1.xml`. */
	stem: string;
	items: readonly ListItem[];
	/** Part names these keys had in the local package. */
	known: Map<string, string>;
	files: Map<string, Uint8Array | string>;
	/** Names in use; extended with the ones placed. */
	taken: Set<string>;
	localZip: JSZip | undefined;
}

const relTarget = (dir: string, path: string): string =>
	dir && path.startsWith(`${dir}/`) ? path.slice(dir.length + 1) : `/${path}`;

/** Writes the list, its relationships and the item parts; returns each item's part name. */
export async function placeList(options: ListOptions): Promise<Map<string, string>> {
	const { files, taken, localZip, items } = options;
	const dir = dirName(options.listPath);
	taken.add(options.listPath).add(relsPathOf(options.listPath));
	const paths = new Map<string, string>();
	for (const item of items) {
		const known = options.known.get(item.key);
		if (known && !taken.has(known) && localZip?.file(known)) {
			paths.set(item.key, known);
			taken.add(known);
		}
	}
	let counter = 0;
	for (const item of items) {
		if (paths.has(item.key)) continue;
		let path: string;
		do path = `${dir ? `${dir}/` : ''}${options.stem}${++counter}.xml`;
		while (taken.has(path) || localZip?.file(path));
		paths.set(item.key, path);
		taken.add(path);
	}
	const others = parseRelationships(
		localZip ? await zipText(localZip, relsPathOf(options.listPath)) : undefined,
	).filter((r) => r.type !== options.relType);
	const usedIds = new Set(others.map((r) => r.id));
	let nextRel = 0;
	const rels: string[] = [];
	const elements: string[] = [];
	for (const item of items) {
		const path = paths.get(item.key)!;
		let id: string;
		do id = `rId${++nextRel}`;
		while (usedIds.has(id));
		usedIds.add(id);
		rels.push(
			`<Relationship Id="${id}" Type="${options.relType}" Target="${escapeAttr(relTarget(dir, path))}"/>`,
		);
		const rel = `<Rel xmlns:r="${REL_NS}" r:id="${id}"/>`;
		elements.push(
			item.xml.includes(REL_MARK)
				? item.xml.replace(REL_MARK, () => rel)
				: item.xml.slice(0, item.xml.lastIndexOf('</')) +
						rel +
						item.xml.slice(item.xml.lastIndexOf('</')),
		);
		files.set(path, item.part);
		if (item.rels !== undefined) files.set(relsPathOf(path), encoder.encode(item.rels));
	}
	files.set(options.listPath, encoder.encode(options.open + elements.join('') + options.close));
	files.set(
		relsPathOf(options.listPath),
		encoder.encode(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
				others.map((r) => r.text).join('') +
				rels.join('') +
				'</Relationships>',
		),
	);
	return paths;
}
