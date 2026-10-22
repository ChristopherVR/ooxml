// Splits a local VSDX into the keyed pieces the shared drawing holds: pages, the shapes (one
// entry per shape, nested shapes included) with their cells, the connects of each page, masters,
// media and every other part verbatim. Parts are split by text span, so the XML a peer shares is
// exactly what the editor wrote.
import type { DrawingKeys } from './keys';
import { relsPathOf } from './package';
import type { Piece } from './pieces';
import { attrOf, scanContainer } from './scan';
import type { ConnectEntry } from './schema';
import { splitContents } from './snapshot-contents';
import { readLayout, readLocalItems, readZip, withMark, zipText } from './snapshot-io';
import { snapshotMasters, snapshotRest } from './snapshot-parts';

export {
	type Layout,
	type LocalItem,
	readLayout,
	readLocalItems,
	readZip,
	zipText,
} from './snapshot-io';

export interface ShapeSnapshot {
	key: string;
	id: string;
	/** Key of the shape this one is nested in. */
	parent?: string;
	open: string;
	pieces: Piece[];
}
export interface ConnectSnapshot {
	key: string;
	entry: ConnectEntry;
}
export interface PageSnapshot {
	key: string;
	id: string;
	back?: string;
	xml: string;
	shell: string;
	rels?: string;
	path: string;
	/** Every shape of the page, parents before the shapes nested in them. */
	shapes: ShapeSnapshot[];
	connects: ConnectSnapshot[];
}
export interface MasterSnapshot {
	key: string;
	id: string;
	xml: string;
	path: string;
	/** The master part with its shapes and connects replaced by markers. */
	part: Uint8Array;
	rels?: string;
	/** Every shape of the master's content, parents before the shapes nested in them. */
	shapes: ShapeSnapshot[];
	connects: ConnectSnapshot[];
}
export interface MediaSnapshot {
	path: string;
	hash: string;
	ct: string;
	d: boolean;
	data: Uint8Array;
}
export interface DrawingSnapshot {
	pagesPath: string;
	pagesOpen: string;
	pagesClose: string;
	pages: PageSnapshot[];
	mastersPath?: string;
	mastersOpen?: string;
	mastersClose?: string;
	masters: MasterSnapshot[];
	media: MediaSnapshot[];
	parts: Map<string, Uint8Array>;
}

export interface SnapshotOptions {
	/** Refuse packages whose shared parts exceed this many bytes. */
	maxPartBytes?: number;
}

export const DEFAULT_MAX_PART_BYTES = 48 * 1024 * 1024;

const SHAPE_ID = /<(?:\w+:)?Shape\s[^>]*?\bID\s*=\s*(["'])(\d+)\1/g;
/** The ids of a shape element and of every shape nested in it, outermost first. */
export function shapeIds(xml: string): string[] {
	return [...xml.matchAll(SHAPE_ID)].map((m) => m[2]!);
}

export async function snapshotDrawing(
	bytes: Uint8Array,
	keys: DrawingKeys,
	options: SnapshotOptions = {},
): Promise<DrawingSnapshot> {
	const zip = await readZip(bytes);
	const layout = await readLayout(zip);
	if (!layout) throw new Error('The package has no Visio pages part.');
	const pagesXml = (await zipText(zip, layout.pagesPath)) ?? '';
	const local = await readLocalItems(zip, layout.pagesPath, pagesXml, 'Pages', 'Page');
	if (!local) throw new Error('The pages part cannot be split into pages.');
	const derived = new Set([layout.pagesPath, relsPathOf(layout.pagesPath)]);
	const idToKey = new Map(local.items.map((p) => [p.id, keys.pageKey(p.id)]));
	const pages: PageSnapshot[] = [];
	for (const page of local.items) {
		const key = idToKey.get(page.id)!;
		derived.add(page.path).add(relsPathOf(page.path));
		const head = page.text.slice(0, scanContainer(page.text, 'Page')!.openEnd);
		const backId = attrOf(head, 'BackPage');
		const content = (await zipText(zip, page.path)) ?? '';
		const rels = await zipText(zip, relsPathOf(page.path));
		const { shell, shapes, connects } = splitContents(content, key, keys);
		keys.pagePaths.set(key, page.path);
		pages.push({
			key,
			id: page.id,
			...(backId !== undefined && idToKey.has(backId) ? { back: idToKey.get(backId)! } : {}),
			xml: withMark(page),
			shell,
			...(rels !== undefined ? { rels } : {}),
			path: page.path,
			shapes,
			connects,
		});
	}

	const { masters, mastersOpen, mastersClose } = await snapshotMasters(zip, layout, keys, derived);
	const { parts, media } = await snapshotRest(
		zip,
		derived,
		options.maxPartBytes ?? DEFAULT_MAX_PART_BYTES,
	);
	return {
		pagesPath: layout.pagesPath,
		pagesOpen: pagesXml.slice(0, local.box.openEnd),
		pagesClose: pagesXml.slice(local.box.closeStart),
		pages,
		...(mastersOpen !== undefined && mastersClose !== undefined && layout.mastersPath
			? { mastersPath: layout.mastersPath, mastersOpen, mastersClose }
			: {}),
		masters,
		media,
		parts,
	};
}
