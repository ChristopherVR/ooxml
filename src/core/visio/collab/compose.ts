// Rebuilds a VSDX from the resolved shared drawing: the shared parts verbatim, a pages part and
// page parts spliced from the shared pages and shapes, and the content types and page
// relationships regenerated for this peer's part names. Parts that stay local (the window state and
// the thumbnail) are kept from the local package.
import JSZip from 'jszip';
import type { DrawingKeys } from './keys';
import {
	CONTENT_TYPES,
	PAGE_TYPE,
	REL_NS,
	decoder,
	dirName,
	encoder,
	fromBase64,
	fingerprint,
	canon,
	isLocalOnlyPart,
	parseRelationships,
	resolveTarget,
	relsPathOf,
	withPageOverrides,
} from './package';
import type { ResolvedDrawing, ResolvedPage } from './read';
import { attrOf, escapeAttr, scanContainer } from './scan';
import { CONNECTS_MARK, REL_MARK, SHAPES_MARK } from './schema';
import { readZip, zipText } from './snapshot';
import { bodyContent, bodyId, connectId, pageContent, pageId, shapeId } from './write';

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

/** The page part with its shapes and connects filled in. */
export function pagePartXml(page: ResolvedPage): string {
	let shell = page.body.shell;
	const root = scanContainer(shell, 'PageContents');
	if (page.shapes.length && !shell.includes(SHAPES_MARK) && root && !root.selfClosing)
		shell =
			shell.slice(0, root.openEnd) + `<Shapes>${SHAPES_MARK}</Shapes>` + shell.slice(root.openEnd);
	const root2 = scanContainer(shell, 'PageContents');
	if (page.connects.length && !shell.includes(CONNECTS_MARK) && root2 && !root2.selfClosing)
		shell =
			shell.slice(0, root2.closeStart) +
			`<Connects>${CONNECTS_MARK}</Connects>` +
			shell.slice(root2.closeStart);
	return shell
		.replace(SHAPES_MARK, () => page.shapes.map((s) => s.xml).join(''))
		.replace(CONNECTS_MARK, () => page.connects.map(connectXml).join(''));
}

function relTarget(pagesDir: string, path: string): string {
	return pagesDir && path.startsWith(`${pagesDir}/`) ? path.slice(pagesDir.length + 1) : `/${path}`;
}

export interface ComposedDrawing {
	bytes: Uint8Array;
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
	for (const [path, b64] of drawing.parts) {
		const data = fromBase64(b64);
		total += data.length;
		if (total > limit) throw new Error('The shared parts exceed the size limit.');
		files.set(path, data);
		prints.set(path, fingerprint(data));
	}
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

	// Part names: a page keeps the part it has locally, a new page gets the next free name.
	const pagesDir = dirName(drawing.pagesPath);
	const taken = new Set<string>([
		...files.keys(),
		drawing.pagesPath,
		relsPathOf(drawing.pagesPath),
	]);
	const paths = new Map<string, string>();
	for (const page of drawing.pages) {
		const known = keys.pagePaths.get(page.key);
		if (known && !taken.has(known) && localZip?.file(known)) {
			paths.set(page.key, known);
			taken.add(known);
		}
	}
	let counter = 0;
	for (const page of drawing.pages) {
		if (paths.has(page.key)) continue;
		let path: string;
		do path = `${pagesDir ? `${pagesDir}/` : ''}page${++counter}.xml`;
		while (taken.has(path) || localZip?.file(path));
		paths.set(page.key, path);
		taken.add(path);
	}

	const others = parseRelationships(
		localZip ? await zipText(localZip, relsPathOf(drawing.pagesPath)) : undefined,
	).filter((r) => r.type !== PAGE_TYPE);
	const usedIds = new Set(others.map((r) => r.id));
	let nextRel = 0;
	const pageRels: string[] = [];
	const pageXml: string[] = [];
	for (const page of drawing.pages) {
		const path = paths.get(page.key)!;
		let id: string;
		do id = `rId${++nextRel}`;
		while (usedIds.has(id));
		usedIds.add(id);
		pageRels.push(
			`<Relationship Id="${id}" Type="${PAGE_TYPE}" Target="${escapeAttr(relTarget(pagesDir, path))}"/>`,
		);
		const rel = `<Rel xmlns:r="${REL_NS}" r:id="${id}"/>`;
		pageXml.push(
			page.xml.includes(REL_MARK)
				? page.xml.replace(REL_MARK, () => rel)
				: page.xml.slice(0, page.xml.lastIndexOf('</')) +
						rel +
						page.xml.slice(page.xml.lastIndexOf('</')),
		);
		files.set(path, encoder.encode(pagePartXml(page)));
		if (page.body.rels !== undefined) files.set(relsPathOf(path), encoder.encode(page.body.rels));
	}
	files.set(
		drawing.pagesPath,
		encoder.encode(drawing.pagesOpen + pageXml.join('') + drawing.pagesClose),
	);
	files.set(
		relsPathOf(drawing.pagesPath),
		encoder.encode(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
				others.map((r) => r.text).join('') +
				pageRels.join('') +
				'</Relationships>',
		),
	);
	files.set(CONTENT_TYPES, encoder.encode(withPageOverrides(contentTypes, [...paths.values()])));

	const zip = new JSZip();
	const names = [...files.keys()].sort((a, b) =>
		a === CONTENT_TYPES ? -1 : b === CONTENT_TYPES ? 1 : 0,
	);
	for (const name of names) zip.file(name, files.get(name)!, { createFolders: false });
	const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'STORE' });

	const commit = (): void => {
		// This peer now knows exactly what the room holds.
		const pageIds = new Map<string, string>();
		const shapeIds = new Map<string, Map<string, string>>();
		const known = new Map<string, string>();
		keys.pagePaths.clear();
		for (const page of drawing.pages) {
			pageIds.set(page.id, page.key);
			keys.pagePaths.set(page.key, paths.get(page.key)!);
			known.set(pageId(page.key), pageContent(page));
			known.set(bodyId(page.key), bodyContent(page.body));
			shapeIds.set(page.key, new Map(page.shapes.map((s) => [s.id, s.key])));
			for (const shape of page.shapes) known.set(shapeId(page.key, shape.key), canon(shape.xml));
			for (const connect of page.connects) known.set(connectId(connect.key), '');
		}
		keys.reset(pageIds, shapeIds);
		keys.known = known;
		keys.partPrints.clear();
		for (const [path, print] of prints) keys.partPrints.set(path, print);
	};
	return { bytes, commit };
}
