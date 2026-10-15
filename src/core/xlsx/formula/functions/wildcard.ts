// A backtracking-free matcher for Excel wildcards (`*`, `?`, `~` escape). It replaces a generated
// regular expression, whose repeated `*` groups backtrack exponentially on long text.

/** A run of text between `*`s: a character, or `null` for `?` (any one character). */
type Segment = (string | null)[];

export interface Wildcard {
	/** Whether the whole text matches the pattern (case-insensitive). */
	test(text: string): boolean;
	/** The zero-based start of the earliest match at or after `from`, or -1. */
	search(text: string, from?: number): number;
}

const lower = (ch: string | undefined): string => (ch ?? '').toLowerCase();
const fold = (text: string): string[] =>
	Array.from({ length: text.length }, (_, i) => lower(text[i]));

interface Parsed {
	segments: Segment[];
	leadingStar: boolean;
	trailingStar: boolean;
}

function parse(pattern: string): Parsed {
	const segments: Segment[] = [];
	let current: Segment = [];
	let leadingStar = false;
	let trailingStar = false;
	const flush = (): void => {
		if (current.length) segments.push(current);
		current = [];
	};
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i] ?? '';
		if (ch === '~' && i + 1 < pattern.length) {
			current.push(lower(pattern[++i]));
			trailingStar = false;
		} else if (ch === '*') {
			if (i === 0) leadingStar = true;
			flush();
			trailingStar = true;
		} else {
			current.push(ch === '?' ? null : lower(ch));
			trailingStar = false;
		}
	}
	flush();
	return { segments, leadingStar, trailingStar };
}

function matchesAt(seg: Segment, text: string[], at: number): boolean {
	if (at < 0 || at + seg.length > text.length) return false;
	for (let i = 0; i < seg.length; i++) {
		const want = seg[i];
		if (want !== null && want !== text[at + i]) return false;
	}
	return true;
}

function find(seg: Segment, text: string[], from: number, limit: number): number {
	for (let at = from; at + seg.length <= limit; at++) if (matchesAt(seg, text, at)) return at;
	return -1;
}

/** Greedy earliest placement of `segments[from..to)` in `text[pos..limit)`; the end, or -1. */
function place(
	segments: Segment[],
	from: number,
	to: number,
	text: string[],
	pos: number,
	limit: number,
): number {
	for (let i = from; i < to; i++) {
		const seg = segments[i] as Segment;
		const at = find(seg, text, pos, limit);
		if (at < 0) return -1;
		pos = at + seg.length;
	}
	return pos;
}

export function compileWildcard(pattern: string): Wildcard {
	const { segments, leadingStar, trailingStar } = parse(pattern);
	const n = segments.length;
	return {
		test(source) {
			const text = fold(source);
			if (n === 0) return leadingStar || trailingStar ? true : text.length === 0;
			const first = segments[0] as Segment;
			const last = segments[n - 1] as Segment;
			if (n === 1 && !leadingStar && !trailingStar) {
				return first.length === text.length && matchesAt(first, text, 0);
			}
			let pos = 0;
			let from = 0;
			let to = n;
			let limit = text.length;
			if (!leadingStar) {
				if (!matchesAt(first, text, 0)) return false;
				pos = first.length;
				from = 1;
			}
			if (!trailingStar) {
				if (to - from < 1) return pos === text.length;
				if (limit - last.length < pos || !matchesAt(last, text, limit - last.length)) return false;
				limit -= last.length;
				to = n - 1;
			}
			return place(segments, from, to, text, pos, limit) >= 0;
		},
		search(source, from = 0) {
			const text = fold(source);
			if (n === 0) return from <= text.length ? from : -1;
			if (leadingStar) return place(segments, 0, n, text, from, text.length) >= 0 ? from : -1;
			const first = segments[0] as Segment;
			const at = find(first, text, from, text.length);
			if (at < 0) return -1;
			return place(segments, 1, n, text, at + first.length, text.length) >= 0 ? at : -1;
		},
	};
}
