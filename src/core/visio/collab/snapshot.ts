// Splits a local VSDX into the keyed pieces the shared drawing holds: pages, the shapes (one
// entry per shape, nested shapes included) with their cells, the connects of each page, masters,
// media and every other part verbatim. Parts are split by text span, so the XML a peer shares is
// exactly what the editor wrote.
import type { DrawingKeys } from './keys';
import { relsPathOf } from './package';
import { type Piece, splitShape } from './pieces';
import { attrOf, scanContainer } from './scan';
import type { ConnectEntry } from './schema';
import { CONNECTS_MARK, SHAPES_MARK } from './schema';
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
	part: Uint8Array;
	rels?: string;
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

/** Replaces the inner content of the first `name` element by `mark`; returns its children. */
function splitList(text: string, name: string, mark: string, child: string) {
	const box = scanContainer(text, name);
	if (!box) return { text, items: [] as string[] };
	const items = box.children.filter((c) => c.name === child).map((c) => text.slice(c.start, c.end));
	if (box.selfClosing) {
		const tag = text.slice(box.start, box.end);
		const qualified = /^<([^\s/>]+)/.exec(tag)?.[1] ?? name;
		const open = tag.replace(/\s*\/>$/, '>');
		return {
			text: text.slice(0, box.start) + open + mark + `</${qualified}>` + text.slice(box.end),
			items,
		};
	}
	return { text: text.slice(0, box.openEnd) + mark + text.slice(box.closeStart), items };
}

function flatten(
	xml: string,
	parent: string | undefined,
	page: string,
	keys: DrawingKeys,
	out: ShapeSnapshot[],
): void {
	const split = splitShape(xml);
	const id = split ? attrOf(split.open, 'ID') : undefined;
	if (!split || id === undefined) return;
	const key = keys.shapeKey(page, id);
	out.push({
		key,
		id,
		...(parent !== undefined ? { parent } : {}),
		open: split.open,
		pieces: split.pieces,
	});
	for (const nested of split.nested) flatten(nested, key, page, keys, out);
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
		const shapesSplit = splitList(content, 'Shapes', SHAPES_MARK, 'Shape');
		const connectsSplit = splitList(shapesSplit.text, 'Connects', CONNECTS_MARK, 'Connect');
		const shapes: ShapeSnapshot[] = [];
		for (const text of shapesSplit.items) flatten(text, undefined, key, keys, shapes);
		const owner = new Map(shapes.map((s) => [s.id, s.key]));
		const connects: ConnectSnapshot[] = [];
		for (const text of connectsSplit.items) {
			const a = (name: string) => attrOf(text, name) ?? '';
			const entry: ConnectEntry = {
				fk: owner.get(a('FromSheet')) ?? '',
				fi: a('FromSheet'),
				tk: owner.get(a('ToSheet')) ?? '',
				ti: a('ToSheet'),
				fc: a('FromCell'),
				tc: a('ToCell'),
				fp: a('FromPart'),
				tp: a('ToPart'),
			};
			connects.push({
				key: `${key}/${entry.fk}.${entry.fi}.${entry.fc}.${entry.fp}>${entry.tk}.${entry.ti}.${entry.tc}.${entry.tp}`,
				entry,
			});
		}
		keys.pagePaths.set(key, page.path);
		pages.push({
			key,
			id: page.id,
			...(backId !== undefined && idToKey.has(backId) ? { back: idToKey.get(backId)! } : {}),
			xml: withMark(page),
			shell: connectsSplit.text,
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
