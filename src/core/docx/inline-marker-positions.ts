// Where a paragraph's position markers (bookmarks, editing permissions, proofing marks) go when
// the paragraph is rebuilt after an edit. Positions count characters; a field marker, picture,
// equation, note reference or break counts as one, so an editor that merges runs of the same
// formatting does not move markers. The paragraph before and after the edit are compared with a
// character diff, so text no edit touched keeps its markers, however many places were edited:
// - a marker inside replaced text moves to the start of the replacement if it opens a range, its
//   end if it closes one;
// - text inserted exactly at a marker stays outside the range it opens or closes;
// - a range from the paragraph's start, or to its end, keeps covering the paragraph.
// An insertion or deletion beside an equal character could have been made on either side of it
// (an "s" typed at the end of a bookmarked "Paris" looks the same inside the bookmark as after
// it): the placement that leaves marked text as it was is used.
import { diffText, type TextChange } from '../text/text-diff';
import type { TextRun } from './model';

const OBJECT = '\uFFFC';

/** Field markers, pictures, equations, notes and breaks are one character; text is its length. */
export function positionText(run: TextRun): string {
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

/** A marker's character position, and whether it opens a range. */
export interface MarkerAt {
	position: number;
	opening: boolean;
	/** An exact end (a wrapper's) never takes in text typed beside it, even at an edge. */
	exact?: boolean;
}

const isSurrogate = (code: number) => code >= 0xd800 && code <= 0xdfff;

/** The same insertion or deletion one character to the left (-1) or right (+1), if equivalent. */
function shifted(
	change: TextChange,
	by: -1 | 1,
	from: string,
	to: string,
	limits: { fromStart: number; toStart: number; fromEnd: number; toEnd: number },
): TextChange | undefined {
	const moved = {
		fromStart: change.fromStart + by,
		fromEnd: change.fromEnd + by,
		toStart: change.toStart + by,
		toEnd: change.toEnd + by,
	};
	if (
		moved.fromStart < limits.fromStart ||
		moved.toStart < limits.toStart ||
		moved.fromEnd > limits.fromEnd ||
		moved.toEnd > limits.toEnd
	)
		return undefined;
	// The character the edit moves across must equal the one it leaves at the other end.
	const insertion = change.fromStart === change.fromEnd;
	const crossed = from.charCodeAt(by < 0 ? moved.fromStart : change.fromStart);
	const left = insertion
		? to.charCodeAt(by < 0 ? change.toEnd - 1 : change.toStart)
		: from.charCodeAt(by < 0 ? change.fromEnd - 1 : change.fromEnd);
	return crossed === left && !isSurrogate(crossed) ? moved : undefined;
}

/** How many markers would lose text from their range if `change` deleted it. */
function swallowed(change: TextChange, markers: readonly MarkerAt[]): number {
	return markers.filter(({ position, opening }) =>
		opening
			? position >= change.fromStart && position < change.fromEnd
			: position > change.fromStart && position <= change.fromEnd,
	).length;
}

/**
 * The placement of an insertion or a deletion of repeated text that keeps marked text as it was:
 * an insertion at the leftmost marker it can reach that keeps it outside that marker's range, a
 * deletion that takes text from the fewest ranges. A replacement has one placement.
 */
function placeChange(
	change: TextChange,
	from: string,
	to: string,
	limits: { fromStart: number; toStart: number; fromEnd: number; toEnd: number },
	markers: readonly MarkerAt[],
): TextChange {
	const insertion = change.fromStart === change.fromEnd;
	if (insertion === (change.toStart === change.toEnd)) return change;
	const placements = [change];
	let left = shifted(change, -1, from, to, limits);
	while (left) {
		placements.unshift(left);
		left = shifted(left, -1, from, to, limits);
	}
	let right = shifted(change, 1, from, to, limits);
	while (right) {
		placements.push(right);
		right = shifted(right, 1, from, to, limits);
	}
	// A marker that keeps a range covering the paragraph's start or end takes in text typed at it,
	// so an insertion placed there would not stay outside its range.
	const keepsOut = ({ position, opening, exact }: MarkerAt) =>
		exact || (opening ? position > 0 : position < from.length);
	if (insertion)
		return (
			placements.find((placement) =>
				markers.some((marker) => marker.position === placement.fromStart && keepsOut(marker)),
			) ?? change
		);
	let best = change;
	for (const placement of placements)
		if (swallowed(placement, markers) < swallowed(best, markers)) best = placement;
	return best;
}

/** Moves each marker from its position in the paragraph `before` an edit to the one `after` it. */
export function moveMarkers<T extends MarkerAt>(
	before: readonly TextRun[],
	after: readonly TextRun[],
	markers: readonly T[],
): T[] {
	const from = before.map(positionText).join('');
	const to = after.map(positionText).join('');
	const changes: TextChange[] = [];
	const found = diffText(from, to);
	for (const [index, change] of found.entries()) {
		const previous = changes.at(-1);
		const next = found[index + 1];
		changes.push(
			placeChange(
				change,
				from,
				to,
				{
					fromStart: previous?.fromEnd ?? 0,
					toStart: previous?.toEnd ?? 0,
					fromEnd: next?.fromStart ?? from.length,
					toEnd: next?.toStart ?? to.length,
				},
				markers,
			),
		);
	}
	const move = ({ position, opening, exact }: MarkerAt): number => {
		// A range from the paragraph's start, or to its end, keeps covering the edited paragraph.
		if (!exact && opening && position === 0) return 0;
		if (!exact && !opening && position === from.length) return to.length;
		let shift = 0;
		for (const [index, change] of changes.entries()) {
			if (position < change.fromStart) break;
			if (position > change.fromEnd) {
				shift = change.toEnd - change.fromEnd;
				continue;
			}
			// Text inserted exactly at a marker stays outside the range it opens or closes, also when
			// the insertion follows a change that ends at the marker.
			const next = changes[index + 1];
			const inserted =
				change.fromStart === change.fromEnd
					? change
					: position === change.fromEnd && next?.fromStart === position && next.fromEnd === position
						? next
						: undefined;
			if (inserted) return opening ? inserted.toEnd : inserted.toStart;
			if (position === change.fromStart) return change.toStart;
			if (position === change.fromEnd) return change.toEnd;
			return opening ? change.toStart : change.toEnd;
		}
		return position + shift;
	};
	return markers.map((marker) => ({ ...marker, position: move(marker) }));
}
