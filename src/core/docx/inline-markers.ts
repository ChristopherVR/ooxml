// Zero-width markers inside a paragraph (bookmarks, editing permissions, proofing marks) are not
// in the run model. When an edited paragraph is rebuilt they are carried to the same character
// position (inline-marker-positions.ts) and inserted there, between runs or inside a split run.
import type { TextRun } from './model';
import type { AnnotationEdge } from './inline-annotations';
import { runPositions } from './inline-marker-positions';
import { symbolCharacter } from './special-characters';
import { getW, isElement, named, type XmlElement, type XmlNode } from './xml';

const POSITION_MARKERS = ['bookmarkStart', 'bookmarkEnd', 'proofErr', 'permStart', 'permEnd'];

/** A zero-width marker the writer carries through a rebuilt paragraph. */
export const isPositionMarker = (element: XmlElement): boolean =>
	POSITION_MARKERS.some((name) => named(element, name));

/** Smart tags and inline custom XML only annotate their runs (see inline-annotations.ts). */
export const isAnnotationWrapper = (element: XmlElement): boolean =>
	named(element, 'smartTag') || named(element, 'customXml');

/** A marker, the number of the paragraph's runs before it, and the container it was read in. */
export interface InlineMarker {
	element: XmlElement;
	boundary: number;
	/** The hyperlink, simple field or content control (`w:sdt`) around the marker. */
	container?: XmlElement;
	/** Set when the marker stands in for one end of a smart tag or custom XML element. */
	annotation?: AnnotationEdge;
}

/** Whether the marker opens a range (`bookmarkStart`, `permStart`, a proofing `...Start`). */
export function opensRange(element: XmlElement): boolean {
	if (named(element, 'proofErr')) return /Start$/.test(getW(element, 'type') ?? '');
	return named(element, 'bookmarkStart') || named(element, 'permStart');
}

/** The paragraph child that holds `node`. */
function topLevel(node: XmlElement, paragraph: XmlElement): XmlElement {
	let current = node;
	while (current.parentNode && current.parentNode !== paragraph)
		current = current.parentNode as XmlElement;
	return current;
}

/** The characters a run element's content stands for, one entry per content element. */
function contentLength(child: XmlElement): number {
	if (named(child, 't') || named(child, 'delText')) return (child.textContent ?? '').length;
	return named(child, 'sym') ? symbolCharacter(getW(child, 'char') ?? '').length || 1 : 1;
}

/**
 * Splits a written text run element at `offset` characters, leaving the first part in place, and
 * returns the second part (inserted right after it), or undefined when `offset` is at its end.
 */
function splitRun(run: XmlElement, offset: number): XmlElement | undefined {
	const children = Array.from(run.childNodes).filter(isElement);
	const properties = children.find((child) => named(child, 'rPr'));
	const second = run.cloneNode(false) as XmlElement;
	if (properties) second.appendChild(properties.cloneNode(true));
	let seen = 0;
	for (const child of children) {
		if (child === properties) continue;
		const length = contentLength(child);
		if (seen >= offset) second.appendChild(child);
		else if (seen + length > offset) {
			// The split falls inside this w:t: the rest of its text moves to the second run.
			const text = child.textContent ?? '';
			const tail = child.cloneNode(false) as XmlElement;
			for (const node of Array.from(child.childNodes)) child.removeChild(node);
			child.appendChild(run.ownerDocument!.createTextNode(text.slice(0, offset - seen)));
			tail.appendChild(run.ownerDocument!.createTextNode(text.slice(offset - seen)));
			for (const part of [child, tail])
				if (/^\s|\s$/.test(part.textContent ?? ''))
					part.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
			second.appendChild(tail);
		}
		seen += length;
	}
	if (second.childNodes.length === (properties ? 1 : 0)) return undefined;
	run.parentNode!.insertBefore(second, run.nextSibling);
	return second;
}

/**
 * Where each run was written: its first and last nodes, the hyperlink, simple field or control
 * content it was written into (none at the paragraph's level), and its `w:r` when a marker may
 * split it.
 */
export interface PlacedRun {
	first?: XmlElement;
	last?: XmlElement;
	parent?: XmlElement;
	splittable?: XmlElement;
}

/** The id that pairs a range's start and end markers (`bookmarkStart`/`End`, `permStart`/`End`). */
function rangeKey(element: XmlElement): string | undefined {
	const id = getW(element, 'id');
	if (id === undefined) return undefined;
	if (named(element, 'bookmarkStart') || named(element, 'bookmarkEnd')) return `bookmark:${id}`;
	if (named(element, 'permStart') || named(element, 'permEnd')) return `perm:${id}`;
	return undefined;
}

/**
 * Keeps every range's end, and a smart tag's or custom XML element's end, at or after its start
 * (an empty range or wrapper typed into stays empty).
 */
export function orderRangeEnds<
	T extends {
		element: XmlElement;
		position: number;
		opening: boolean;
		annotation?: AnnotationEdge;
	},
>(markers: T[]): T[] {
	const keyOf = (marker: T) => marker.annotation?.wrapper ?? rangeKey(marker.element);
	const starts = new Map<XmlElement | string, number>();
	for (const marker of markers) {
		const key = keyOf(marker);
		if (key && marker.opening) starts.set(key, marker.position);
	}
	return markers.map((marker) => {
		const key = keyOf(marker);
		const start = !marker.opening && key ? starts.get(key) : undefined;
		return start !== undefined && marker.position < start ? { ...marker, position: start } : marker;
	});
}

/**
 * Inserts each marker at its character position: between the runs there, or inside a text run,
 * which is split in two. A position inside a run that cannot be split (a tracked change) moves to
 * the run's start for a marker that opens a range, its end for one that closes it. A marker read
 * inside a hyperlink, simple field or content control stays in what was written in its place
 * (`containers` maps each source container to those elements) when a run beside the marker is
 * there; any other marker goes between the paragraph's children, or between two runs of one
 * container.
 */
export function placeMarkers(
	paragraph: XmlElement,
	markers: readonly {
		element: XmlElement;
		position: number;
		opening: boolean;
		container?: XmlElement;
		annotation?: AnnotationEdge;
	}[],
	runs: readonly TextRun[],
	placed: PlacedRun[],
	containers: ReadonlyMap<XmlElement, ReadonlySet<XmlElement>>,
): void {
	const positions = runPositions(runs);
	// The parts a split run now has, each with the position of its first character.
	const pieces = new Map<number, { element: XmlElement; start: number }[]>();
	const insideRun = (index: number, position: number, element: XmlElement): boolean => {
		const run = placed[index]?.splittable;
		if (!run) return false;
		const parts = pieces.get(index) ?? [{ element: run, start: positions[index]! }];
		pieces.set(index, parts);
		let at = parts.length - 1;
		while (at > 0 && parts[at]!.start > position) at--;
		const part = parts[at]!;
		if (part.start === position) {
			part.element.parentNode!.insertBefore(element, part.element);
			return true;
		}
		const second = splitRun(part.element, position - part.start);
		if (!second) return false;
		parts.splice(at + 1, 0, { element: second, start: position });
		// A marker at the run's end now goes after its second part.
		const written = placed[index];
		if (written?.last === part.element) written.last = second;
		second.parentNode!.insertBefore(element, second);
		return true;
	};
	// A marker placed after a node goes after the markers already placed there, which come before
	// it in the source.
	const done = new Set<XmlNode>();
	const after = (node: XmlElement, element: XmlElement) => {
		let anchor: XmlNode = node;
		while (anchor.nextSibling && done.has(anchor.nextSibling)) anchor = anchor.nextSibling;
		anchor.parentNode!.insertBefore(element, anchor.nextSibling);
	};
	const put = (marker: (typeof markers)[number]) => {
		let boundary = marker.opening
			? positions.indexOf(marker.position)
			: positions.lastIndexOf(marker.position);
		if (boundary < 0) {
			const inside = positions.findIndex(
				(start, index) => start < marker.position && marker.position < positions[index + 1]!,
			);
			if (insideRun(inside, marker.position, marker.element)) return;
			boundary = marker.opening ? inside : inside + 1;
		}
		const previous = placed[boundary - 1];
		const next = placed[boundary];
		const home = marker.container && containers.get(marker.container);
		// A smart tag's or custom XML element's edges stay beside the text it wraps, in whatever
		// container that text is now written: its start before the first run, its end after the last.
		const side = marker.annotation?.edge;
		if (side === 'start' && next?.first?.parentNode)
			return void next.first.parentNode.insertBefore(marker.element, next.first);
		if (side === 'end' && previous?.last) return after(previous.last, marker.element);
		if (home && previous?.last && previous.parent && home.has(previous.parent))
			return after(previous.last, marker.element);
		if (home && next?.first?.parentNode && next.parent && home.has(next.parent))
			return void next.first.parentNode.insertBefore(marker.element, next.first);
		if (!next?.first?.parentNode) return void paragraph.appendChild(marker.element);
		// Between two runs written into one container the marker stays inside it.
		const holder = topLevel(next.first, paragraph);
		const shared =
			previous?.last && holder !== next.first && topLevel(previous.last, paragraph) === holder;
		const before = shared ? next.first : holder;
		before.parentNode!.insertBefore(marker.element, before);
	};
	// An empty smart tag or custom XML element keeps its ends together: its end goes where its
	// start went, after the markers placed there since (the ones it held).
	const starts = new Map<XmlElement, { parent: XmlNode; next: XmlNode | null; position: number }>();
	for (const marker of markers) {
		marker.element.parentNode?.removeChild(marker.element);
		const annotation = marker.annotation;
		const start = annotation?.edge === 'end' ? starts.get(annotation.wrapper) : undefined;
		if (
			start?.position === marker.position &&
			(!start.next || start.next.parentNode === start.parent)
		) {
			start.parent.insertBefore(marker.element, start.next);
			done.add(marker.element);
			continue;
		}
		put(marker);
		done.add(marker.element);
		if (annotation?.edge === 'start' && marker.element.parentNode)
			starts.set(annotation.wrapper, {
				parent: marker.element.parentNode,
				next: marker.element.nextSibling,
				position: marker.position,
			});
	}
}
