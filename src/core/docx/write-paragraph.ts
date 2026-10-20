// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Paragraph } from './model';
import { children, first, makeW, named, WORD_NS, type XmlDocument, type XmlElement } from './xml';
import {
	markerPositions,
	opensRange,
	orderRangeEnds,
	placeMarkers,
	runPositions,
} from './inline-markers';
import { writeParagraphProperties } from './write-paragraph-properties';
import { orderParagraphProperties, writeTabStops } from './tab-stops';
import { continuationKey } from './comment-spans';
import type { CommentSpans } from './write-ranges';
import { reconcileBookmarks } from './bookmarks';
import { writeNumberingProperties } from './numbering-write';
import { writeParagraphFormatRevision, writeParagraphMarkRevision } from './write-revisions';
import { paragraphJustification } from './paragraph-alignment';
import { runHasUnknownProperties } from './write-run-validation';
import { buildInlineContent } from './write-inline';
import { collectInlineSlots, replaceableInlineChildren } from './inline-slots';
import type { RelationshipAllocator } from './relationship-allocator';
import { parseDirectParagraphProperties } from './paragraph-properties';
import { parsePropertiesSnapshot } from './revision-properties';
import { samePictureSource } from './inline-source';

export function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
}

function rejectUnsafeRunSegmentation(
	paragraph: Paragraph,
	base: Paragraph | undefined,
	oldRuns: XmlElement[],
): void {
	if (!base || !oldRuns.some(runHasUnknownProperties)) return;
	// Pictures are matched by their media source, so splitting surrounding text does not move
	// their opaque properties to a different run. Retained opaque text still needs its own basis.
	if (
		oldRuns.length === base.runs.length &&
		oldRuns.every((node, index) => {
			if (!runHasUnknownProperties(node)) return true;
			const image = base.runs[index]?.image;
			if (!image) return false;
			// A drawing without a media part (a shape, text box or chart) is the same drawing when
			// its model is unchanged; a changed one is rebuilt from its own source element.
			const remaining = paragraph.runs.filter((run) =>
				image.partName
					? samePictureSource(run.image, image)
					: run.image !== undefined && JSON.stringify(run.image) === JSON.stringify(image),
			);
			return (
				!remaining.length ||
				remaining.every((run) => run.restoredRunPropertiesXml || run.sourceRunPropertiesXml)
			);
		})
	)
		return;
	// Imported opaque properties or a rejected format revision supply a complete XML basis.
	// The writer imports that basis before applying modeled edits, independently of source slots.
	if (
		paragraph.runs.every(
			(run) =>
				!run.image && !run.equation && (run.restoredRunPropertiesXml || run.sourceRunPropertiesXml),
		)
	)
		return;
	const sameText =
		paragraph.runs.map((run) => run.text).join('') === base.runs.map((run) => run.text).join('');
	const sameBoundaries =
		paragraph.runs.length === base.runs.length &&
		paragraph.runs.every((run, index) => run.text === base.runs[index]?.text);
	if (sameBoundaries || (!sameText && paragraph.runs.length === base.runs.length)) return;
	throw new Error(
		`Cannot edit paragraph ${paragraph.id}: changing run boundaries could drop unsupported run properties. The original DOCX package remains unchanged.`,
	);
}

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
	let pPr = first(node, 'pPr');
	let propertyBase = base;
	const restoring =
		paragraph.restoredParagraphPropertiesXml &&
		(!base || paragraph.restoredParagraphPropertiesXml !== base.restoredParagraphPropertiesXml);
	const propertiesXml =
		paragraph.restoredParagraphPropertiesXml ?? paragraph.sourceParagraphPropertiesXml;
	if (propertiesXml && (restoring || paragraph.sourceParagraphPropertiesXml || !base)) {
		const restored = parsePropertiesSnapshot(propertiesXml, 'pPr');
		if (pPr) node.removeChild(pPr);
		pPr = doc.importNode(restored, true) as XmlElement;
		node.insertBefore(pPr, node.firstChild);
		propertyBase = {
			type: 'paragraph',
			id: paragraph.id,
			runs: paragraph.runs,
			...parseDirectParagraphProperties(restored),
		};
	}
	if (!pPr) {
		pPr = makeW(doc, 'pPr');
		node.insertBefore(pPr, node.firstChild);
	}
	if (
		!propertyBase ||
		paragraph.align !== propertyBase.align ||
		paragraph.justification !== propertyBase.justification ||
		paragraph.direction !== propertyBase.direction
	) {
		removeChildren(pPr, 'jc');
		const jc = paragraphJustification(paragraph);
		if (jc) {
			const align = makeW(doc, 'jc');
			setAttribute(align, 'val', jc);
			pPr.appendChild(align);
		}
	}
	if (!propertyBase || paragraph.style !== propertyBase.style) {
		removeChildren(pPr, 'pStyle');
		if (paragraph.style) {
			const style = makeW(doc, 'pStyle');
			setAttribute(style, 'val', paragraph.style);
			pPr.appendChild(style);
		}
	}
	writeParagraphProperties(doc, pPr, paragraph, propertyBase);
	writeNumberingProperties(doc, pPr, paragraph, propertyBase);
	writeParagraphMarkRevision(doc, pPr, paragraph, restoring ? undefined : base);
	writeParagraphFormatRevision(doc, pPr, paragraph.formatRevision);
	writeTabStops(doc, pPr, paragraph, propertyBase);
	orderParagraphProperties(pPr);
	const { slots, markers } = layout;
	rejectUnsafeRunSegmentation(
		paragraph,
		base,
		slots.flatMap((slot) => (slot.element ? [slot.element] : [])),
	);
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
	const positionOf = base
		? markerPositions(base.runs, paragraph.runs)
		: (_position: number, opening: boolean) => (opening ? 0 : runPositions(paragraph.runs).at(-1)!);
	const carried = orderRangeEnds(
		markers
			.filter((marker) => !isBookmark(marker.element) || kept.has(marker.element))
			.map((marker) => {
				const opening = opensRange(marker.element);
				return {
					element: marker.element,
					position: positionOf(sourcePositions[marker.boundary] ?? 0, opening),
					opening,
					topLevel: marker.element.parentNode === node,
				};
			}),
	);
	for (const marker of markers) marker.element.parentNode?.removeChild(marker.element);
	for (const oldNode of replaceableInlineChildren(node)) node.removeChild(oldNode);
	const { nodes, placed } = buildInlineContent(
		doc,
		paragraph.runs,
		base?.runs,
		slots,
		allocator,
		ranges,
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
	placeMarkers(node, carried, paragraph.runs, placed);
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
