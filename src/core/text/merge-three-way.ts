import { diffText } from './text-diff';

export type TextMergeResult = { status: 'merged'; text: string } | { status: 'unresolved' };
interface Edit {
	start: number;
	end: number;
	text: string;
}

/** Combine non-overlapping changes against one shared ancestor, without inventing edit history. */
export function mergeTextThreeWay(base: string, local: string, remote: string): TextMergeResult {
	if (local === remote || base === remote) return { status: 'merged', text: local };
	if (base === local) return { status: 'merged', text: remote };
	if (Math.max(base.length, local.length, remote.length) > 1_000_000)
		return { status: 'unresolved' };
	const changes = (value: string): Edit[] =>
		diffText(base, value).map((change) => ({
			start: change.fromStart,
			end: change.fromEnd,
			text: value.slice(change.toStart, change.toEnd),
		}));
	const edits = changes(local);
	for (const next of changes(remote)) {
		if (
			edits.some(
				(edit) => edit.start === next.start && edit.end === next.end && edit.text === next.text,
			)
		)
			continue;
		if (
			edits.some((edit) => {
				// Concurrent insertions both survive; a deterministic text order breaks snapshot ties.
				if (edit.start === edit.end && next.start === next.end) return false;
				if (edit.start === edit.end) return edit.start > next.start && edit.start < next.end;
				if (next.start === next.end) return next.start > edit.start && next.start < edit.end;
				return edit.start < next.end && next.start < edit.end;
			})
		)
			return { status: 'unresolved' };
		edits.push(next);
	}
	edits.sort(
		(a, b) =>
			a.start - b.start || a.end - b.end || (a.text < b.text ? -1 : a.text > b.text ? 1 : 0),
	);
	let text = '',
		cursor = 0;
	for (const edit of edits) {
		text += base.slice(cursor, edit.start) + edit.text;
		cursor = edit.end;
	}
	return { status: 'merged', text: text + base.slice(cursor) };
}
