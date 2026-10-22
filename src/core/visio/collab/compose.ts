// Rebuilds a VSDX from the resolved shared drawing: the shared parts verbatim, media from their
// blobs, a pages part and page parts spliced from the shared pages and shapes, a masters part and
// master parts from the master entries, and the content types and relationships regenerated for
// this peer's part names. Parts that stay local (the window state and the thumbnail) are kept from
// the local package. Page parts that did not change since the last rebuild are not serialised
// again (the cache in `keys.cache`); the package itself is still zipped and parsed whole.
import type JSZip from 'jszip';
import {
	bodyContent,
	bodyId,
	cellId,
	connectId,
	entryContent,
	masterContent,
	masterId,
	masterScope,
	mediaId,
	pageContent,
	pageId,
	shapeId,
	shellContent,
} from './baseline';
import { type ListItem, placeList } from './compose-lists';
import type { DrawingKeys } from './keys';
import {
	CONTENT_TYPES,
	MASTER_CONTENT_TYPE,
	MASTER_TYPE,
	PAGE_TYPE,
	decoder,
	dirName,
	encoder,
	ensureType,
	fingerprint,
	fromBase64,
	isLocalOnlyPart,
	resolveTarget,
	withOverrides,
	withPageOverrides,
} from './package';
import type { ResolvedDrawing, ResolvedMaster, ResolvedPage, ResolvedShape } from './read';
import { attrOf, escapeAttr, scanContainer } from './scan';
import { CONNECTS_MARK, SHAPES_MARK } from './schema';
import { readZip, zipText } from './snapshot';
import { writeStoredZip } from './zip-store';

export interface ComposeOptions {
	/** Refuse shared parts that decode to more than this many bytes in total. */
	maxPartBytes?: number;
}

const DEFAULT_CONTENT_TYPES =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';

function connectXml(item: ResolvedPage['connects'][number]): string {
	const attrs: [string, string][] = [
		['FromSheet', item.from],
		['FromCell', item.entry.fc],
		['FromPart', item.entry.fp],
		['ToSheet', item.to],
		['ToCell', item.entry.tc],
		['ToPart', item.entry.tp],
	];
	return `<Connect ${attrs
		.filter(([, v]) => v !== '')
		.map(([k, v]) => `${k}="${escapeAttr(v)}"`)
		.join(' ')}/>`;
}

/** A content part (`PageContents`, `MasterContents`) with its shapes and connects filled in. */
function contentsXml(
	shell: string,
	root: string,
	shapesXml: string,
	hasShapes: boolean,
	connects: ResolvedPage['connects'],
): string {
	const outer = scanContainer(shell, root);
	if (hasShapes && !shell.includes(SHAPES_MARK) && outer && !outer.selfClosing)
		shell =
			shell.slice(0, outer.openEnd) +
			`<Shapes>${SHAPES_MARK}</Shapes>` +
			shell.slice(outer.openEnd);
	const outer2 = scanContainer(shell, root);
	if (connects.length && !shell.includes(CONNECTS_MARK) && outer2 && !outer2.selfClosing)
		shell =
			shell.slice(0, outer2.closeStart) +
			`<Connects>${CONNECTS_MARK}</Connects>` +
			shell.slice(outer2.closeStart);
	return shell
		.replace(SHAPES_MARK, () => shapesXml)
		.replace(CONNECTS_MARK, () => connects.map(connectXml).join(''));
}

/** The page part with its shapes and connects filled in. */
export const pagePartXml = (page: ResolvedPage): string =>
	contentsXml(
		page.body.shell,
		'PageContents',
		page.shapesXml,
		page.shapes.length > 0,
		page.connects,
	);

/** The master part with its shapes and connects filled in. */
export const masterPartXml = (master: ResolvedMaster): string =>
	contentsXml(
		decoder.decode(fromBase64(master.part)),
		'MasterContents',
		master.shapesXml,
		master.shapes.length > 0,
		master.connects,
	);

export interface ComposedDrawing {
	bytes: Uint8Array;
	/** The resolved drawing's signature: equal signatures build equal packages. */
	signature: string;
	/** Records what the room held as this peer's baseline; call it once the editor adopted `bytes`. */
	commit: () => void;
}

export async function composeDrawing(
	drawing: ResolvedDrawing,
	local: Uint8Array | undefined,
	keys: DrawingKeys,
	options: ComposeOptions = {},
): Promise<ComposedDrawing> {
	const localZip: JSZip | undefined = local ? await readZip(local) : undefined;
	const files = new Map<string, Uint8Array | string>();
	const prints = new Map<string, string>();
	let total = 0;
	const limit = options.maxPartBytes ?? 48 * 1024 * 1024;
	const cache = keys.cache.parts;
	const decode = (path: string, b64: string): Uint8Array => {
		const hit = cache.get(path);
		const data = hit && hit.b64 === b64 ? hit.data : fromBase64(b64);
		if (!hit || hit.b64 !== b64) cache.set(path, { b64, data });
		total += data.length;
		if (total > limit) throw new Error('The shared parts exceed the size limit.');
		return data;
	};
	for (const [path, b64] of drawing.parts) {
		const data = decode(path, b64);
		files.set(path, data);
		prints.set(path, fingerprint(data));
	}
	for (const path of [...cache.keys()])
		if (!drawing.parts.has(path) && !path.startsWith('media:')) cache.delete(path);
	if (localZip)
		for (const path of Object.keys(localZip.files)) {
			const file = localZip.files[path]!;
			if (!file.dir && isLocalOnlyPart(path) && !files.has(path))
				files.set(path, await file.async('uint8array'));
		}
	let contentTypes =
		files.get(CONTENT_TYPES) !== undefined
			? decoder.decode(files.get(CONTENT_TYPES) as Uint8Array)
			: ((localZip && (await zipText(localZip, CONTENT_TYPES))) ?? DEFAULT_CONTENT_TYPES);
	for (const item of drawing.media) {
		files.set(item.path, decode(`media:${item.path}`, item.data));
		contentTypes = ensureType(contentTypes, item.path, item.declared);
	}

	// A part that stays local and that this peer does not have (a joiner's first package) must not
	// stay referenced.
	for (const [path, data] of [...files]) {
		if (!path.endsWith('.rels') || typeof data === 'string') continue;
		const source = path.replace(/(^|\/)_rels\/([^/]+)\.rels$/, '$1$2');
		const base = dirName(source);
		const text = decoder.decode(data);
		const next = text.replace(/<(?:\w+:)?Relationship\b[^>]*>/g, (tag) => {
			const target = attrOf(tag, 'Target');
			if (target === undefined || attrOf(tag, 'TargetMode') === 'External') return tag;
			const resolved = resolveTarget(base, target);
			if (!isLocalOnlyPart(resolved) || files.has(resolved)) return tag;
			contentTypes = contentTypes.replace(/<(?:\w+:)?Override\b[^>]*>/g, (o) =>
				attrOf(o, 'PartName') === `/${resolved}` ? '' : o,
			);
			return '';
		});
		if (next !== text) files.set(path, encoder.encode(next));
	}

	const taken = new Set<string>(files.keys());
	const pageCache = keys.cache.pageParts;
	const pageItems: ListItem[] = drawing.pages.map((page) => {
		const hit = pageCache.get(page.key);
		const bytes = hit && hit.sig === page.signature ? hit.bytes : encoder.encode(pagePartXml(page));
		pageCache.set(page.key, { sig: page.signature, bytes });
		return {
			key: page.key,
			xml: page.xml,
			part: bytes,
			...(page.body.rels !== undefined ? { rels: page.body.rels } : {}),
		};
	});
	for (const key of [...pageCache.keys()])
		if (!drawing.pages.some((p) => p.key === key)) pageCache.delete(key);
	const pagePaths = await placeList({
		listPath: drawing.pagesPath,
		open: drawing.pagesOpen,
		close: drawing.pagesClose,
		relType: PAGE_TYPE,
		stem: 'page',
		items: pageItems,
		known: keys.pagePaths,
		files,
		taken,
		localZip,
	});
	contentTypes = withPageOverrides(contentTypes, [...pagePaths.values()]);
	let masterPaths = new Map<string, string>();
	if (drawing.mastersPath !== undefined) {
		masterPaths = await placeList({
			listPath: drawing.mastersPath,
			open: drawing.mastersOpen ?? '<Masters>',
			close: drawing.mastersClose ?? '</Masters>',
			relType: MASTER_TYPE,
			stem: 'master',
			items: drawing.masters.map((m) => ({
				key: m.key,
				xml: m.xml,
				part: encoder.encode(masterPartXml(m)),
				...(m.rels !== undefined ? { rels: m.rels } : {}),
			})),
			known: keys.masterPaths,
			files,
			taken,
			localZip,
		});
		contentTypes = withOverrides(contentTypes, MASTER_CONTENT_TYPE, [...masterPaths.values()]);
	}
	files.set(CONTENT_TYPES, encoder.encode(contentTypes));

	// Stored entries: unchanged parts keep their checksum between rebuilds (see zip-store.ts).
	const names = [...files.keys()].sort((a, b) =>
		a === CONTENT_TYPES ? -1 : b === CONTENT_TYPES ? 1 : 0,
	);
	const bytes = writeStoredZip(
		names.map((name) => {
			const data = files.get(name)!;
			return [name, typeof data === 'string' ? encoder.encode(data) : data] as const;
		}),
		keys.cache.crcs,
	);

	const commit = (): void => {
		// This peer now knows exactly what the room held.
		const pageIds = new Map<string, string>();
		const shapeIds = new Map<string, Map<string, string>>();
		const masterIds = new Map<string, string>();
		const known = new Map<string, string>();
		const recordShapes = (
			scope: string,
			shapes: readonly ResolvedShape[],
			connects: ResolvedPage['connects'],
		): void => {
			shapeIds.set(scope, new Map(shapes.map((s) => [s.id, s.key])));
			for (const shape of shapes) {
				known.set(shapeId(scope, shape.key), shellContent(shape));
				for (const [path, entry] of shape.entries)
					known.set(cellId(scope, shape.key, path), entryContent(entry));
			}
			for (const connect of connects) known.set(connectId(connect.key), '');
		};
		keys.pagePaths.clear();
		keys.masterPaths.clear();
		keys.mediaKeys.clear();
		for (const page of drawing.pages) {
			pageIds.set(page.id, page.key);
			keys.pagePaths.set(page.key, pagePaths.get(page.key)!);
			known.set(pageId(page.key), pageContent(page));
			known.set(bodyId(page.key), bodyContent(page.body));
			recordShapes(page.key, page.shapes, page.connects);
		}
		for (const master of drawing.masters) {
			masterIds.set(master.id, master.key);
			recordShapes(masterScope(master.key), master.shapes, master.connects);
			const path = masterPaths.get(master.key);
			if (path) keys.masterPaths.set(master.key, path);
			known.set(
				masterId(master.key),
				masterContent({ ...master, print: fingerprint(fromBase64(master.part)) }),
			);
		}
		for (const item of drawing.media) {
			known.set(mediaId(item.key), '');
			keys.mediaKeys.set(item.path, item.key);
		}
		keys.reset(pageIds, shapeIds, masterIds);
		keys.known = known;
		keys.partPrints.clear();
		for (const [path, print] of prints) keys.partPrints.set(path, print);
	};
	return { bytes, signature: drawing.signature, commit };
}
