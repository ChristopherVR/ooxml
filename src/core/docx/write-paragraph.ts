// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Paragraph } from './model';
import { makeW, named, type XmlDocument, type XmlElement } from './xml';
import { opensRange, orderRangeEnds, placeMarkers } from './inline-markers';
import { moveMarkers, runPositions } from './inline-marker-positions';
import { restoreAnnotations } from './inline-annotations';
import { continuationKey } from './comment-spans';
import type { CommentSpans } from './write-ranges';
import { reconcileBookmarks } from './bookmarks';
import { changedPermissionMarkers } from './permission-markers';
import { buildInlineContent } from './write-inline';
import { containerSources } from './inline-containers';
import { followContainers } from './inline-container-markers';
import { followFormFieldBookmarks } from './inline-form-field-bookmarks';
import { collectInlineSlots, replaceableInlineChildren } from './inline-slots';
import type { RelationshipAllocator } from './relationship-allocator';
import { runsToWriteAnew } from './run-segmentation';
import { writeParagraphMark } from './write-paragraph-mark';

export { setAttribute } from './write-paragraph-mark';

export function writeParagraphImpl(
	doc: XmlDocument,
	paragraph: Paragraph,
	node: XmlElement,
	base: Paragraph | undefined,
	allocator: RelationshipAllocator | undefined,
	spans?: CommentSpans,
): XmlElement {
	const continuation = spans?.next.get(paragraph.id);
	const ranges = spans && {
		...(continuation ?? { before: new Set<string>(), after: new Set<string>() }),
		rangeIds: spans.rangeIds,
	};
	const sameRanges =
		continuationKey(continuation) === continuationKey(spans?.original.get(paragraph.id));
	if (base && sameRanges && JSON.stringify(paragraph) === JSON.stringify(base)) return node;
	const layout = collectInlineSlots(node);
	if (!layout)
		throw new Error(
			`Cannot edit paragraph ${paragraph.id}: it contains inline OOXML that this editor cannot safely relocate. The original DOCX package remains unchanged.`,
		);
	const { pPr } = writeParagraphMark(doc, paragraph, node, base);
	const { slots, markers } = layout;
	const anew = runsToWriteAnew(paragraph, base, slots);
	// Bookmarks, permissions and proofing marks keep their character positions (inline-markers.ts).
	// Bookmark names removed from the model lose their markers; names added bracket the paragraph.
	const isBookmark = (element: XmlElement) =>
		named(element, 'bookmarkStart') || named(element, 'bookmarkEnd');
	const sourceBookmarks = markers.map((marker) => marker.element).filter(isBookmark);
	const { starts, ends } = reconcileBookmarks(
		doc,
		sourceBookmarks.filter((element) => named(element, 'bookmarkStart')),
		sourceBookmarks.filter((element) => named(element, 'bookmarkEnd')),
		paragraph.bookmarks ?? [],
		base?.bookmarks ?? [],
		() => spans?.nextId() ?? '0',
	);
	const source = new Set(sourceBookmarks);
	const kept = new Set([...starts, ...ends]);
	const sourcePositions = runPositions(base?.runs ?? []);
	const permissions = changedPermissionMarkers(
		doc,
		paragraph,
		base,
		markers.map((marker) => marker.element),
		runPositions(paragraph.runs).at(-1)!,
	);
	const read = markers
		.filter((marker) => !isBookmark(marker.element) || kept.has(marker.element))
		.filter((marker) => !permissions?.drop.has(marker.element))
		.map((marker) => {
			const position = sourcePositions[marker.boundary] ?? 0;
			const annotation = marker.annotation;
			return {
				element: marker.element,
				position,
				source: position,
				opening: annotation ? annotation.edge === 'start' : opensRange(marker.element),
				...(marker.container && { container: marker.container }),
				...(marker.revision && { revision: marker.revision }),
				...(annotation && { annotation, exact: true }),
			};
		});
	const end = runPositions(paragraph.runs).at(-1)!;
	const sources = containerSources(paragraph.runs, base?.runs, slots);
	const carried = orderRangeEnds([
		...(base
			? followFormFieldBookmarks(
					followContainers(
						moveMarkers(base.runs, paragraph.runs, read),
						sources,
						base.runs,
						paragraph.runs,
					),
					base.runs,
					paragraph.runs,
				)
			: read.map((marker) => ({ ...marker, position: marker.opening ? 0 : end }))),
		...(permissions?.added ?? []),
	]);
	for (const marker of markers) marker.element.parentNode?.removeChild(marker.element);
	for (const oldNode of replaceableInlineChildren(node)) node.removeChild(oldNode);
	const { nodes, placed, containers } = buildInlineContent(
		doc,
		paragraph.runs,
		base?.runs,
		slots,
		allocator,
		ranges,
		sources,
		anew,
	);
	let anchor: XmlElement = pPr;
	for (const bookmark of starts.filter((element) => !source.has(element))) {
		node.insertBefore(bookmark, anchor.nextSibling);
		anchor = bookmark;
	}
	for (const run of nodes) {
		node.insertBefore(run, anchor.nextSibling);
		anchor = run;
	}
	placeMarkers(node, carried, paragraph.runs, placed, containers);
	// Smart tags and inline custom XML go back around text the edit left as it was.
	restoreAnnotations(carried, base?.runs ?? [], paragraph.runs);
	for (const bookmark of ends.filter((element) => !source.has(element))) node.appendChild(bookmark);
	if (!pPr.childNodes.length) node.removeChild(pPr);
	return node;
}

export function createParagraph(
	doc: XmlDocument,
	paragraph: Paragraph,
	allocator: RelationshipAllocator | undefined,
	spans?: CommentSpans,
): XmlElement {
	const node = makeW(doc, 'p');
	return writeParagraphImpl(doc, paragraph, node, undefined, allocator, spans);
}
