// Zero-width markers inside a paragraph (bookmarks, editing permissions, proofing marks) are not
// in the run model. When an edited paragraph is rebuilt they are carried to the same character
// position: text before and after the edit keeps its markers, a marker inside the edited text
// moves to its start if it opens a range or its end if it closes one, and a range from the
// paragraph's start or to its end keeps covering the paragraph. Positions count
// characters, so an editor that merges runs of the same formatting does not move markers; a
// field marker, picture or note reference counts as one character.
import type { TextRun } from './model';
import { symbolCharacter } from './special-characters';
import { getW, isElement, named, type XmlElement } from './xml';

const POSITION_MARKERS = ['bookmarkStart', 'bookmarkEnd', 'proofErr', 'permStart', 'permEnd'];
const OBJECT = '\uFFFC';

/** A zero-width marker the writer carries through a rebuilt paragraph. */
export const isPositionMarker = (element: XmlElement): boolean =>
	POSITION_MARKERS.some((name) => named(element, name));

/** Smart tags and inline custom XML only annotate their runs; a rebuilt paragraph keeps the runs. */
export const isAnnotationWrapper = (element: XmlElement): boolean =>
	named(element, 'smartTag') || named(element, 'customXml');

/** A marker and the number of the paragraph's runs before it. */
export interface InlineMarker {
	element: XmlElement;
	boundary: number;
}

/** Whether the marker opens a range (`bookmarkStart`, `permStart`, a proofing `...Start`). */
export function opensRange(element: XmlElement): boolean {
	if (named(element, 'proofErr')) return /Start$/.test(getW(element, 'type') ?? '');
	return named(element, 'bookmarkStart') || named(element, 'permStart');
}

/** Field markers, pictures, equations, notes and breaks are one character; text is its length. */
function positionText(run: TextRun): string {
	return run.fieldChar ||
		run.fieldCode !== undefined ||
		run.image ||
		run.equation ||
		run.noteReference ||
		run.noteMark ||
		run.noteSeparator ||
		run.break
		? OBJECT
		: run.text;
}

/** The character position of each run boundary (0 to `runs.length`). */
export function runPositions(runs: readonly TextRun[]): number[] {
	const positions = [0];
	for (const run of runs) positions.push(positions.at(-1)! + positionText(run).length);
	return positions;
}

const isHighSurrogate = (code: number) => code >= 0xd800 && code <= 0xdbff;
const isLowSurrogate = (code: number) => code >= 0xdc00 && code <= 0xdfff;

/** Maps a character position in the paragraph `before` an edit onto the paragraph `after` it. */
export function markerPositions(
	before: readonly TextRun[],
	after: readonly TextRun[],
): (position: number, opening: boolean) => number {
	const from = before.map(positionText).join('');
	const to = after.map(positionText).join('');
	let prefix = 0;
	while (prefix < from.length && prefix < to.length && from[prefix] === to[prefix]) prefix++;
	// Never end the unchanged stretch inside a surrogate pair: no marker may split a character.
	if (prefix > 0 && isHighSurrogate(from.charCodeAt(prefix - 1))) prefix--;
	let suffix = 0;
	while (
		suffix < from.length - prefix &&
		suffix < to.length - prefix &&
		from[from.length - 1 - suffix] === to[to.length - 1 - suffix]
	)
		suffix++;
	if (suffix > 0 && isLowSurrogate(from.charCodeAt(from.length - suffix))) suffix--;
	const delta = to.length - from.length;
	return (position, opening) => {
		// A range from the paragraph's start, or to its end, keeps covering the edited paragraph.
		if (opening && position === 0) return 0;
		if (!opening && position === from.length) return to.length;
		const leading = position <= prefix;
		const trailing = position >= from.length - suffix;
		// Text inserted exactly at a marker stays outside the range it opens or closes.
		if (leading && trailing) return opening ? position + delta : position;
		if (leading) return position;
		if (trailing) return position + delta;
		return opening ? prefix : to.length - suffix;
	};
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

/** Where each run was written: its first node, and its `w:r` when a marker may split it. */
export interface PlacedRun {
	first?: XmlElement;
	splittable?: XmlElement;
}

/**
 * Inserts each marker at its character position: between the runs there, or inside a text run,
 * which is split in two. A position inside a run that cannot be split (a tracked change) moves to
 * the run's start for a marker that opens a range, its end for one that closes it. A marker that
 * was a direct child of the paragraph stays one when its run starts a hyperlink or content control.
 */
/** The id that pairs a range's start and end markers (`bookmarkStart`/`End`, `permStart`/`End`). */
function rangeKey(element: XmlElement): string | undefined {
	const id = getW(element, 'id');
	if (id === undefined) return undefined;
	if (named(element, 'bookmarkStart') || named(element, 'bookmarkEnd')) return `bookmark:${id}`;
	if (named(element, 'permStart') || named(element, 'permEnd')) return `perm:${id}`;
	return undefined;
}

/** Keeps every range's end at or after its start (an empty range typed into stays empty). */
export function orderRangeEnds<
	T extends { element: XmlElement; position: number; opening: boolean },
>(markers: T[]): T[] {
	const starts = new Map<string, number>();
	for (const marker of markers) {
		const key = rangeKey(marker.element);
		if (key && marker.opening) starts.set(key, marker.position);
	}
	return markers.map((marker) => {
		const start = !marker.opening && starts.get(rangeKey(marker.element) ?? '');
		return start !== undefined && start !== false && marker.position < start
			? { ...marker, position: start }
			: marker;
	});
}

export function placeMarkers(
	paragraph: XmlElement,
	markers: readonly {
		element: XmlElement;
		position: number;
		opening: boolean;
		topLevel: boolean;
	}[],
	runs: readonly TextRun[],
	placed: PlacedRun[],
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
		second.parentNode!.insertBefore(element, second);
		return true;
	};
	for (const marker of markers) {
		marker.element.parentNode?.removeChild(marker.element);
		let boundary = marker.opening
			? positions.indexOf(marker.position)
			: positions.lastIndexOf(marker.position);
		if (boundary < 0) {
			const inside = positions.findIndex(
				(start, index) => start < marker.position && marker.position < positions[index + 1]!,
			);
			if (insideRun(inside, marker.position, marker.element)) continue;
			boundary = marker.opening ? inside : inside + 1;
		}
		const first = placed[boundary]?.first;
		if (!first?.parentNode) {
			paragraph.appendChild(marker.element);
			continue;
		}
		const holder = topLevel(first, paragraph);
		const previous = placed[boundary - 1]?.first;
		const startsHolder = holder === first || !previous || topLevel(previous, paragraph) !== holder;
		const before = marker.topLevel && startsHolder ? holder : first;
		before.parentNode!.insertBefore(marker.element, before);
	}
}
