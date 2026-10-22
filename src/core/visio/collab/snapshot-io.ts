// Reading a local package for the snapshot: the zip, its layout and the items of a list part.
import JSZip from 'jszip';
import { dirName, parseRelationships, relsPathOf, resolveTarget } from './package';
import { attrOf, scanContainer } from './scan';
import { REL_MARK } from './schema';

export const readZip = (bytes: Uint8Array): Promise<JSZip> => JSZip.loadAsync(bytes);
export const zipText = async (zip: JSZip, path: string): Promise<string | undefined> =>
	zip.file(path)?.async('string');

export interface Layout {
	documentPath: string;
	pagesPath: string;
	mastersPath?: string;
}

/** The document, pages and masters parts of a package, through its relationships. */
export async function readLayout(zip: JSZip): Promise<Layout | undefined> {
	const root = parseRelationships(await zipText(zip, '_rels/.rels')).find((r) =>
		r.type.endsWith('/document'),
	);
	if (!root) return undefined;
	const documentPath = resolveTarget('', root.target);
	const rels = parseRelationships(await zipText(zip, relsPathOf(documentPath)));
	const pages = rels.find((r) => r.type.endsWith('/pages'));
	const masters = rels.find((r) => r.type.endsWith('/masters'));
	const dir = dirName(documentPath);
	return pages
		? {
				documentPath,
				pagesPath: resolveTarget(dir, pages.target),
				...(masters ? { mastersPath: resolveTarget(dir, masters.target) } : {}),
			}
		: undefined;
}

export interface LocalItem {
	id: string;
	text: string;
	path: string;
	rel: { start: number; end: number };
}

/** The `child` elements of the `container` of a list part (`Pages`, `Masters`) with their parts. */
export async function readLocalItems(
	zip: JSZip,
	listPath: string,
	xml: string,
	container: string,
	child: string,
) {
	const box = scanContainer(xml, container);
	if (!box) return undefined;
	const rels = new Map(
		parseRelationships(await zipText(zip, relsPathOf(listPath))).map((r) => [r.id, r]),
	);
	const items: LocalItem[] = [];
	for (const span of box.children) {
		if (span.name !== child) continue;
		const text = xml.slice(span.start, span.end);
		const inner = scanContainer(text, child);
		const relSpan = inner?.children.find((c) => c.name === 'Rel');
		if (!inner || !relSpan) return undefined;
		const relTag = text.slice(relSpan.start, relSpan.openEnd);
		const rid = /\s(?:\w+:)?id\s*=\s*(["'])([^"']+)\1/.exec(relTag)?.[2];
		const target = rid ? rels.get(rid)?.target : undefined;
		const id = attrOf(text.slice(0, inner.openEnd), 'ID');
		if (target === undefined || id === undefined) return undefined;
		items.push({
			id,
			text,
			path: resolveTarget(dirName(listPath), target),
			rel: { start: relSpan.start, end: relSpan.end },
		});
	}
	return { box, items };
}

export const withMark = (item: LocalItem): string =>
	item.text.slice(0, item.rel.start) + REL_MARK + item.text.slice(item.rel.end);
