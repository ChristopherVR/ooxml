// Smart tags (`w:smartTag`) and inline custom XML (`w:customXml`) only annotate the runs they
// wrap; the run model keeps their text. When an edited paragraph is rebuilt, each wrapper is put
// back around the same text if that text did not change, with its own properties. The markup
// around text that changed is dropped, and a load warning says so.
import { isSdt } from './content-control';
import { positionText } from './inline-marker-positions';
import { isPositionMarker } from './inline-markers';
import type { TextRun } from './model';
import { isRevisionWrapperElement } from './write-revisions';
import { isElement, named, WORD_NS, type XmlElement, type XmlNode } from './xml';

/** One end of a smart tag or inline custom XML element, carried like a position marker. */
export interface AnnotationEdge {
	wrapper: XmlElement;
	edge: 'start' | 'end';
}

/** A zero-width stand-in for one end of `wrapper` while the paragraph is rebuilt. */
export const annotationEdge = (wrapper: XmlElement): XmlElement =>
	wrapper.ownerDocument!.createElementNS(WORD_NS, 'w:annotationEdge');

/** Markers and stand-ins, which a stand-in can be moved across without changing any text. */
const zeroWidth = (node: XmlNode): boolean =>
	!isElement(node) || isPositionMarker(node) || named(node, 'annotationEdge');

function within(node: XmlNode, ancestor: XmlNode): boolean {
	for (let at: XmlNode | null = node; at; at = at.parentNode) if (at === ancestor) return true;
	return false;
}

/**
 * When the text a wrapper marked now starts or ends inside a hyperlink or simple field (a link
 * was added over part of it), or inside one of the tracked changes the writer gives each run of a
 * change, moves the stand-in from the container's first or last place to just before or after the
 * container, which the restored wrapper then holds. A content control is not left: a wrapper
 * around one could not be edited again.
 */
function liftOut(opening: XmlElement, closing: XmlElement): void {
	const container = (edge: XmlElement): XmlElement | undefined => {
		const parent = edge.parentNode;
		return parent &&
			isElement(parent) &&
			(named(parent, 'hyperlink') || named(parent, 'fldSimple') || isRevisionWrapperElement(parent))
			? parent
			: undefined;
	};
	// The markers and stand-ins beside the moved one go with it, in order, so an enclosing
	// wrapper's own stand-in stays outside the restored one.
	for (let parent = container(opening); parent && !within(closing, parent);) {
		const moving: XmlNode[] = [opening];
		for (let node = opening.previousSibling; node; node = node.previousSibling) {
			if (!zeroWidth(node)) return;
			moving.unshift(node);
		}
		for (const node of moving) parent.parentNode!.insertBefore(node, parent);
		parent = container(opening);
	}
	for (let parent = container(closing); parent && !within(opening, parent);) {
		const moving: XmlNode[] = [closing];
		for (let node = closing.nextSibling; node; node = node.nextSibling) {
			if (!zeroWidth(node)) return;
			moving.push(node);
		}
		const after = parent.nextSibling;
		for (const node of moving) parent.parentNode!.insertBefore(node, after);
		parent = container(closing);
	}
}

/**
 * Wraps the content between each annotation's two stand-ins in a copy of its source wrapper when
 * the text it marked is unchanged and both stand-ins share a parent, then removes the stand-ins.
 * `source` is a stand-in's position before the edit, `position` after it.
 */
export function restoreAnnotations(
	markers: readonly {
		element: XmlElement;
		position: number;
		source: number;
		annotation?: AnnotationEdge;
	}[],
	before: readonly TextRun[],
	after: readonly TextRun[],
): void {
	const from = before.map(positionText).join('');
	const to = after.map(positionText).join('');
	const starts = new Map<XmlElement, (typeof markers)[number]>();
	// Inner wrappers end first, so an outer one wraps their restored copies.
	for (const marker of markers) {
		const annotation = marker.annotation;
		if (!annotation) continue;
		if (annotation.edge === 'start') {
			starts.set(annotation.wrapper, marker);
			continue;
		}
		const start = starts.get(annotation.wrapper);
		const opening = start?.element;
		const closing = marker.element;
		if (opening && opening.parentNode !== closing.parentNode) liftOut(opening, closing);
		// What lies between the two stand-ins, if the closing one follows the opening one.
		const between: NonNullable<XmlElement['nextSibling']>[] = [];
		let node = opening?.nextSibling ?? null;
		while (node && node !== closing) {
			between.push(node);
			node = node.nextSibling;
		}
		// A wrapper is not put around a content control: this editor reads controls only at the
		// paragraph's level, so the paragraph could not be edited again.
		if (
			start &&
			opening?.parentNode &&
			node === closing &&
			!between.some((item) => isElement(item) && isSdt(item)) &&
			from.slice(start.source, marker.source) === to.slice(start.position, marker.position)
		) {
			const copy = annotation.wrapper.cloneNode(false) as XmlElement;
			for (const child of Array.from(annotation.wrapper.childNodes))
				if (isElement(child) && (named(child, 'smartTagPr') || named(child, 'customXmlPr')))
					copy.appendChild(child.cloneNode(true));
			opening.parentNode.insertBefore(copy, opening);
			for (const item of between) copy.appendChild(item);
		}
		opening?.parentNode?.removeChild(opening);
		closing.parentNode?.removeChild(closing);
	}
}
