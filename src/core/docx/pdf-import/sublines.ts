/**
 * Pieces of a line: the fragments a table cell or a page column takes from a line become a line of
 * their own, with the column-gap tab that joined them to their neighbours removed.
 */
import type { Fragment, Line } from './lines';

const blank = (fragment: Fragment) => !fragment.text.trim();

/** The ink extent of a fragment: trailing spaces and tabs (which producers draw) do not count. */
export function inkRight(fragment: Fragment): number {
	const kept = fragment.text.trimEnd().length / Math.max(1, fragment.text.length);
	return fragment.x + fragment.width * kept;
}

/** The width of a line's first word, points (what a line before it lacked room for if it wrapped). */
export function firstWordWidth(line: Line): number {
	const fragment = line.fragments.find((item) => !blank(item));
	if (!fragment) return 0;
	const word = /^\S+/.exec(fragment.text.trimStart())?.[0] ?? '';
	return (fragment.width * word.length) / Math.max(1, fragment.text.length);
}

/** A line made of some of `line`'s fragments, or undefined when they hold no text. */
export function subLine(line: Line, fragments: Fragment[]): Line | undefined {
	const kept = fragments.map((fragment) => ({
		...fragment,
		text: fragment.text.replace(/\t/g, ' '),
	}));
	while (kept.length && blank(kept[0]!)) kept.shift();
	while (kept.length && blank(kept[kept.length - 1]!)) kept.pop();
	if (!kept.length) return undefined;
	const last = kept[kept.length - 1]!;
	last.text = last.text.trimEnd();
	kept[0]!.text = kept[0]!.text.trimStart();
	return {
		...line,
		fragments: kept,
		left: kept[0]!.x,
		right: inkRight(last),
		text: kept.map((fragment) => fragment.text).join(''),
		bold: kept.every((fragment) => fragment.font.bold || blank(fragment)),
	};
}

/** The ink intervals of a line, left to right, joined where they are closer than `join` points. */
export function inkIntervals(line: Line, join: number): [number, number][] {
	const out: [number, number][] = [];
	for (const fragment of line.fragments) {
		if (blank(fragment)) continue;
		const start = fragment.x;
		const end = inkRight(fragment);
		const last = out[out.length - 1];
		if (last && start - last[1] < join) last[1] = Math.max(last[1], end);
		else out.push([start, end]);
	}
	return out;
}

/** Splits a line at the `cuts` (x positions, ascending): one piece (maybe undefined) per band. */
export function splitLine(line: Line, cuts: number[]): (Line | undefined)[] {
	const groups: Fragment[][] = Array.from({ length: cuts.length + 1 }, () => []);
	let band = 0;
	for (const fragment of line.fragments) {
		const middle = blank(fragment) ? fragment.x : fragment.x + fragment.width / 2;
		if (!blank(fragment)) {
			band = 0;
			while (band < cuts.length && middle >= cuts[band]!) band++;
		}
		groups[band]!.push(fragment);
	}
	return groups.map((fragments) => subLine(line, fragments));
}
