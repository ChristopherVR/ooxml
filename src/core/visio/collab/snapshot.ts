// Splits a local VSDX into the keyed pieces the shared drawing holds: pages, the shapes and
// connects of each page, and every other part verbatim. Page parts are split by text span, so the
// shape XML is exactly what the editor wrote.
import JSZip from 'jszip';
import type { DrawingKeys } from './keys';
import {
	CONTENT_TYPES,
	decoder,
	dirName,
	encoder,
	isLocalOnlyPart,
	parseRelationships,
	relsPathOf,
	resolveTarget,
	stripPageOverrides,
} from './package';
import { attrOf, scanContainer } from './scan';
import type { ConnectEntry } from './schema';
import { CONNECTS_MARK, REL_MARK, SHAPES_MARK } from './schema';

export interface ShapeSnapshot {
	key: string;
	id: string;
	xml: string;
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
	shapes: ShapeSnapshot[];
	connects: ConnectSnapshot[];
}
export interface DrawingSnapshot {
	pagesPath: string;
	pagesOpen: string;
	pagesClose: string;
	pages: PageSnapshot[];
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

export const readZip = (bytes: Uint8Array): Promise<JSZip> => JSZip.loadAsync(bytes);
export const zipText = async (zip: JSZip, path: string): Promise<string | undefined> =>
	zip.file(path)?.async('string');

export interface Layout {
	documentPath: string;
	pagesPath: string;
}

/** The document and pages parts of a package, through its relationships. */
export async function readLayout(zip: JSZip): Promise<Layout | undefined> {
	const root = parseRelationships(await zipText(zip, '_rels/.rels')).find((r) =>
		r.type.endsWith('/document'),
	);
	if (!root) return undefined;
	const documentPath = resolveTarget('', root.target);
	const rels = parseRelationships(await zipText(zip, relsPathOf(documentPath)));
	const pages = rels.find((r) => r.type.endsWith('/pages'));
	return pages
		? { documentPath, pagesPath: resolveTarget(dirName(documentPath), pages.target) }
		: undefined;
}

export interface LocalPage {
	id: string;
	text: string;
	path: string;
	rel: { start: number; end: number };
}

/** The `Page` elements of a pages part with the part each one points to. */
export async function readLocalPages(zip: JSZip, layout: Layout, pagesXml: string) {
	const box = scanContainer(pagesXml, 'Pages');
	if (!box) return undefined;
	const rels = new Map(
		parseRelationships(await zipText(zip, relsPathOf(layout.pagesPath))).map((r) => [r.id, r]),
	);
	const pages: LocalPage[] = [];
	for (const span of box.children) {
		if (span.name !== 'Page') continue;
		const text = pagesXml.slice(span.start, span.end);
		const inner = scanContainer(text, 'Page');
		const relSpan = inner?.children.find((c) => c.name === 'Rel');
		if (!inner || !relSpan) return undefined;
		const relTag = text.slice(relSpan.start, relSpan.openEnd);
		const rid = /\s(?:\w+:)?id\s*=\s*(["'])([^"']+)\1/.exec(relTag)?.[2];
		const target = rid ? rels.get(rid)?.target : undefined;
		const id = attrOf(text.slice(0, inner.openEnd), 'ID');
		if (target === undefined || id === undefined) return undefined;
		pages.push({
			id,
			text,
			path: resolveTarget(dirName(layout.pagesPath), target),
			rel: { start: relSpan.start, end: relSpan.end },
		});
	}
	return { box, pages };
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
	const local = await readLocalPages(zip, layout, pagesXml);
	if (!local) throw new Error('The pages part cannot be split into pages.');
	const derived = new Set([layout.pagesPath, relsPathOf(layout.pagesPath)]);
	const idToKey = new Map(local.pages.map((p) => [p.id, keys.pageKey(p.id)]));
	const pages: PageSnapshot[] = [];
	for (const page of local.pages) {
		const key = idToKey.get(page.id)!;
		derived.add(page.path).add(relsPathOf(page.path));
		const head = page.text.slice(0, scanContainer(page.text, 'Page')!.openEnd);
		const backId = attrOf(head, 'BackPage');
		const xml = page.text.slice(0, page.rel.start) + REL_MARK + page.text.slice(page.rel.end);
		const content = (await zipText(zip, page.path)) ?? '';
		const rels = await zipText(zip, relsPathOf(page.path));
		const shapesSplit = splitList(content, 'Shapes', SHAPES_MARK, 'Shape');
		const connectsSplit = splitList(shapesSplit.text, 'Connects', CONNECTS_MARK, 'Connect');
		const owner = new Map<string, string>();
		const shapes: ShapeSnapshot[] = [];
		for (const text of shapesSplit.items) {
			const ids = shapeIds(text);
			const id = ids[0];
			if (id === undefined) continue;
			const shapeKey = keys.shapeKey(key, id);
			for (const nested of ids) owner.set(nested, shapeKey);
			shapes.push({ key: shapeKey, id, xml: text });
		}
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
			xml,
			shell: connectsSplit.text,
			...(rels !== undefined ? { rels } : {}),
			path: page.path,
			shapes,
			connects,
		});
	}
	const parts = new Map<string, Uint8Array>();
	let total = 0;
	const limit = options.maxPartBytes ?? DEFAULT_MAX_PART_BYTES;
	for (const path of Object.keys(zip.files)) {
		const file = zip.files[path]!;
		if (file.dir || derived.has(path) || isLocalOnlyPart(path)) continue;
		let data = await file.async('uint8array');
		if (path === CONTENT_TYPES) data = encoder.encode(stripPageOverrides(decoder.decode(data)));
		total += data.length;
		if (total > limit) throw new Error('The package parts exceed the shared size limit.');
		parts.set(path, data);
	}
	return {
		pagesPath: layout.pagesPath,
		pagesOpen: pagesXml.slice(0, local.box.openEnd),
		pagesClose: pagesXml.slice(local.box.closeStart),
		pages,
		parts,
	};
}
