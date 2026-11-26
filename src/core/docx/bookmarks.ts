// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Bookmarks are modeled at paragraph granularity: the model exposes which names start in a
// paragraph, not their exact character range. New names bracket the whole paragraph.
import {
	children,
	getW,
	isElement,
	makeW,
	named,
	WORD_NS,
	type XmlDocument,
	type XmlElement,
} from './xml';
import type { Block, Paragraph } from './model';
import { mapBlockParagraphs } from './document-paragraphs';

const endpoints = new WeakMap<Paragraph, { starts: [string, string][]; ends: string[] }>();

/** Bookmark names starting directly inside this paragraph (`w:bookmarkStart/@w:name`, direct children only). */
export function paragraphBookmarkNames(node: XmlElement, paragraph: Paragraph): string[] {
	const starts: [string, string][] = [];
	const names = children(node, 'bookmarkStart').flatMap((element) => {
		const name = getW(element, 'name');
		const id = getW(element, 'id');
		if (name && id) starts.push([id, name]);
		return name ?? [];
	});
	const ends = children(node, 'bookmarkEnd').flatMap((element) => getW(element, 'id') ?? []);
	if (starts.length || ends.length) endpoints.set(paragraph, { starts, ends });
	return names;
}

/** The closing paragraph of existing cross-paragraph bookmarks, by bookmark name. */
export function sourceBookmarkEnds(blocks: Block[]): [string, string][] {
	const starts = new Map<string, { name: string; paragraphId: string }>();
	const ranges: [string, string][] = [];
	mapBlockParagraphs(blocks, (paragraph) => {
		const markers = endpoints.get(paragraph);
		if (!markers) return paragraph;
		for (const [id, name] of markers.starts) starts.set(id, { name, paragraphId: paragraph.id });
		for (const id of markers.ends) {
			const start = starts.get(id);
			if (start && start.paragraphId !== paragraph.id) ranges.push([start.name, paragraph.id]);
			starts.delete(id);
		}
		return paragraph;
	});
	return ranges;
}

/** Clones a paragraph's direct `w:bookmarkStart`/`w:bookmarkEnd` children so they can be re-inserted after a rewrite. */
export function captureBookmarkNodes(node: XmlElement): {
	starts: XmlElement[];
	ends: XmlElement[];
} {
	const starts: XmlElement[] = [];
	const ends: XmlElement[] = [];
	for (const child of Array.from(node.childNodes)) {
		if (!isElement(child)) continue;
		if (named(child, 'bookmarkStart')) starts.push(child.cloneNode(true) as XmlElement);
		else if (named(child, 'bookmarkEnd')) ends.push(child.cloneNode(true) as XmlElement);
	}
	return { starts, ends };
}

/**
 * A rewritten paragraph's bookmark markers: names added to the model get a start and end around
 * the whole paragraph; names removed from it (present in `base`) lose their markers in this
 * paragraph. Other markers are kept as they were.
 */
export function reconcileBookmarks(
	doc: XmlDocument,
	starts: XmlElement[],
	ends: XmlElement[],
	names: readonly string[],
	baseNames: readonly string[],
	nextId: () => string,
): { starts: XmlElement[]; ends: XmlElement[] } {
	const removedIds = new Set<string>();
	const keptStarts = starts.filter((start) => {
		const name = getW(start, 'name') ?? '';
		const removed = baseNames.includes(name) && !names.includes(name);
		if (removed) removedIds.add(getW(start, 'id') ?? '');
		return !removed;
	});
	const keptEnds = ends.filter((end) => !removedIds.has(getW(end, 'id') ?? ''));
	const present = new Set(keptStarts.map((start) => getW(start, 'name')));
	for (const name of names) {
		if (present.has(name)) continue;
		const id = nextId();
		const start = makeW(doc, 'bookmarkStart');
		start.setAttributeNS(WORD_NS, 'w:id', id);
		start.setAttributeNS(WORD_NS, 'w:name', name);
		const end = makeW(doc, 'bookmarkEnd');
		end.setAttributeNS(WORD_NS, 'w:id', id);
		keptStarts.push(start);
		keptEnds.push(end);
	}
	return { starts: keptStarts, ends: keptEnds };
}
