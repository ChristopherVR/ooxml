// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word names a legacy form field's bookmark after the field (`w:ffData/w:name`) and keeps it around
// the field, from its begin marker to its end marker. Every field reads as the same characters in
// the paragraph's text diff, so swapping two fields, or moving one, changes nothing the diff can
// see: a rebuilt paragraph puts such a bookmark back around its field wherever the field went.
import { runPositions, type MarkerAt } from './inline-marker-positions';
import type { TextRun } from './model';
import { getW, named, type XmlElement } from './xml';

interface Span {
	start: number;
	end: number;
}

/** Where each named form field is: from its begin marker to after its end marker. */
function formFieldSpans(runs: readonly TextRun[]): Map<string, Span[]> {
	const positions = runPositions(runs);
	const spans = new Map<string, Span[]>();
	runs.forEach((run, index) => {
		const name = run.fieldChar === 'begin' ? run.formField?.name : undefined;
		if (!name) return;
		let depth = 0;
		for (let at = index; at < runs.length; at++) {
			const kind = runs[at]!.fieldChar;
			if (kind === 'begin') depth++;
			else if (kind === 'end' && --depth === 0) {
				const span = { start: positions[index]!, end: positions[at + 1]! };
				spans.set(name, [...(spans.get(name) ?? []), span]);
				return;
			}
		}
	});
	return spans;
}

/**
 * Puts each form field's own bookmark (named after the field, and around exactly that field before
 * the edit) around the field again, once the paragraph's markers were moved (`markers`, with their
 * `source` positions in `before`). A bookmark whose field is gone, or is no longer the only one of
 * its name, stays where it was put.
 */
export function followFormFieldBookmarks<
	T extends MarkerAt & { element: XmlElement; source: number },
>(markers: readonly T[], before: readonly TextRun[], after: readonly TextRun[]): T[] {
	const was = formFieldSpans(before);
	if (!was.size) return [...markers];
	const now = formFieldSpans(after);
	const ends = new Map<string, T>();
	for (const marker of markers) {
		const id = named(marker.element, 'bookmarkEnd') ? getW(marker.element, 'id') : undefined;
		if (id !== undefined) ends.set(id, marker);
	}
	const placed = new Map<T, number>();
	for (const start of markers) {
		if (!named(start.element, 'bookmarkStart')) continue;
		const name = getW(start.element, 'name');
		const id = getW(start.element, 'id');
		const end = id === undefined ? undefined : ends.get(id);
		const [from, other] = (name && was.get(name)) || [];
		const [to, another] = (name && now.get(name)) || [];
		if (!end || !from || other || !to || another) continue;
		if (start.source !== from.start || end.source !== from.end) continue;
		if (start.position === to.start && end.position === to.end) continue;
		placed.set(start, to.start);
		placed.set(end, to.end);
	}
	if (!placed.size) return [...markers];
	// Markers at one place are written in the order they come here: a bookmark put back around its
	// field closes before, and opens after, every other marker there, so it holds only the field.
	const moved = (marker: T) => ({ ...marker, position: placed.get(marker)! });
	return [
		...markers.filter((marker) => placed.has(marker) && !marker.opening).map(moved),
		...markers.filter((marker) => !placed.has(marker)),
		...markers.filter((marker) => placed.has(marker) && marker.opening).map(moved),
	];
}
