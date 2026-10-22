// The changes that turn one string into another: a shortest edit script from Myers' O(ND) diff
// ("An O(ND) Difference Algorithm and Its Variations", 1986), over UTF-16 code units. No change
// starts or ends inside a surrogate pair. Past its edit limit the diff compares words instead, so
// text left alone between far-apart edits stays unchanged.

/** `from.slice(fromStart, fromEnd)` became `to.slice(toStart, toEnd)`. */
export interface TextChange {
	fromStart: number;
	fromEnd: number;
	toStart: number;
	toEnd: number;
}

const isLowSurrogate = (text: string, at: number): boolean => {
	const code = text.charCodeAt(at);
	return code >= 0xdc00 && code <= 0xdfff;
};

/**
 * The insertions and deletions of a shortest edit script between two sequences of lengths `n`
 * and `m` whose items `same(x, y)` compares, as `[x, y, inserts]`: the edit at item `x` of the
 * first (a deletion) or item `y` of the second (an insertion). Undefined past `maxEdits` edits.
 */
function shortestEdit(
	n: number,
	m: number,
	same: (x: number, y: number) => boolean,
	maxEdits: number,
): [number, number, boolean][] | undefined {
	const max = Math.min(n + m, maxEdits);
	const offset = max + 1;
	// `v[offset + k]` is the furthest x reached on diagonal k (x - y = k) with d edits.
	const v = new Int32Array(2 * max + 3);
	const trace: Int32Array[] = [];
	for (let d = 0; d <= max; d++) {
		for (let k = -d; k <= d; k += 2) {
			const down = k === -d || (k !== d && v[offset + k - 1]! < v[offset + k + 1]!);
			let x = down ? v[offset + k + 1]! : v[offset + k - 1]! + 1;
			let y = x - k;
			while (x < n && y < m && same(x, y)) {
				x++;
				y++;
			}
			v[offset + k] = x;
			if (x >= n && y >= m) return backtrack(trace, n, m, d);
		}
		trace.push(v.slice(offset - d, offset + d + 1));
	}
	return undefined;
}

/** Walks the furthest-reaching paths back from `(n, m)`, `d` edits away from the start. */
function backtrack(
	trace: readonly Int32Array[],
	n: number,
	m: number,
	d: number,
): [number, number, boolean][] {
	const edits: [number, number, boolean][] = [];
	let x = n;
	let y = m;
	for (let step = d; step > 0; step--) {
		const previous = trace[step - 1]!;
		const at = (k: number) => previous[k + step - 1]!;
		const k = x - y;
		const down = k === -step || (k !== step && at(k - 1) < at(k + 1));
		const before = down ? k + 1 : k - 1;
		const startX = at(before);
		const startY = startX - before;
		edits.push(down ? [startX, startY, true] : [startX, startY, false]);
		x = startX;
		y = startY;
	}
	return edits.reverse();
}

/** Single edits gathered into changes, each a stretch of deleted and inserted items. */
function changesOf(edits: readonly [number, number, boolean][]): TextChange[] {
	const changes: TextChange[] = [];
	for (const [x, y, inserts] of edits) {
		const last = changes.at(-1);
		if (last && last.fromEnd === x && last.toEnd === y) {
			if (inserts) last.toEnd++;
			else last.fromEnd++;
		} else
			changes.push({
				fromStart: x,
				fromEnd: inserts ? x : x + 1,
				toStart: y,
				toEnd: inserts ? y + 1 : y,
			});
	}
	return changes;
}

/** The characters' changes from `a` to `b`, or undefined past `maxEdits` edits. */
function characterChanges(a: string, b: string, maxEdits: number): TextChange[] | undefined {
	const edits = shortestEdit(
		a.length,
		b.length,
		(x, y) => a.charCodeAt(x) === b.charCodeAt(y),
		maxEdits,
	);
	return edits && changesOf(edits);
}

// Words, runs of white space, and single other characters (whole code points).
const TOKEN = /[\p{L}\p{N}\p{M}]+|\s+|[\s\S]/gu;

/** Each token of `text` as a number shared by equal tokens, and where each token starts. */
function tokensOf(text: string, ids: Map<string, number>): { tokens: number[]; starts: number[] } {
	const tokens: number[] = [];
	const starts: number[] = [];
	for (const match of text.matchAll(TOKEN)) {
		let id = ids.get(match[0]);
		if (id === undefined) ids.set(match[0], (id = ids.size));
		tokens.push(id);
		starts.push(match.index);
	}
	starts.push(text.length);
	return { tokens, starts };
}

/**
 * The changes from `a` to `b` found by comparing their tokens, each changed stretch compared
 * again by its characters while the edit limit lasts; undefined when the tokens differ by more
 * than `maxEdits` too.
 */
function tokenChanges(a: string, b: string, maxEdits: number): TextChange[] | undefined {
	const ids = new Map<string, number>();
	const from = tokensOf(a, ids);
	const to = tokensOf(b, ids);
	const edits = shortestEdit(
		from.tokens.length,
		to.tokens.length,
		(x, y) => from.tokens[x] === to.tokens[y],
		maxEdits,
	);
	if (!edits) return undefined;
	let budget = maxEdits;
	return changesOf(edits).flatMap((stretch) => {
		const change = {
			fromStart: from.starts[stretch.fromStart]!,
			fromEnd: from.starts[stretch.fromEnd]!,
			toStart: to.starts[stretch.toStart]!,
			toEnd: to.starts[stretch.toEnd]!,
		};
		const inner =
			change.fromEnd > change.fromStart && change.toEnd > change.toStart
				? characterChanges(
						a.slice(change.fromStart, change.fromEnd),
						b.slice(change.toStart, change.toEnd),
						budget,
					)
				: undefined;
		if (!inner) return [change];
		for (const found of inner)
			budget -= found.fromEnd - found.fromStart + found.toEnd - found.toStart;
		return inner.map((found) => ({
			fromStart: found.fromStart + change.fromStart,
			fromEnd: found.fromEnd + change.fromStart,
			toStart: found.toStart + change.toStart,
			toEnd: found.toEnd + change.toStart,
		}));
	});
}

/**
 * The changes from `from` to `to`, in order and separated by unchanged text. When the strings
 * differ by more than `maxEdits` inserted or deleted code units, their words are compared instead
 * (the characters of each changed stretch again, while the limit lasts); past `maxEdits` changed
 * words too, everything between their common start and end is reported as one change.
 */
export function diffText(from: string, to: string, maxEdits = 500): TextChange[] {
	let start = 0;
	while (
		start < from.length &&
		start < to.length &&
		from.charCodeAt(start) === to.charCodeAt(start)
	)
		start++;
	let end = 0;
	while (
		end < from.length - start &&
		end < to.length - start &&
		from.charCodeAt(from.length - 1 - end) === to.charCodeAt(to.length - 1 - end)
	)
		end++;
	const a = from.slice(start, from.length - end);
	const b = to.slice(start, to.length - end);
	if (!a.length && !b.length) return [];
	const found =
		a.length && b.length
			? (characterChanges(a, b, maxEdits) ?? tokenChanges(a, b, maxEdits))
			: undefined;
	const changes = found ?? [{ fromStart: 0, fromEnd: a.length, toStart: 0, toEnd: b.length }];
	const result: TextChange[] = [];
	for (const change of changes) {
		const placed = {
			fromStart: change.fromStart + start,
			fromEnd: change.fromEnd + start,
			toStart: change.toStart + start,
			toEnd: change.toEnd + start,
		};
		// Widen a change that would split a surrogate pair; the unit beside it is unchanged.
		if (
			placed.fromStart > 0 &&
			placed.toStart > 0 &&
			(isLowSurrogate(from, placed.fromStart) || isLowSurrogate(to, placed.toStart))
		) {
			placed.fromStart--;
			placed.toStart--;
		}
		if (isLowSurrogate(from, placed.fromEnd) || isLowSurrogate(to, placed.toEnd)) {
			placed.fromEnd++;
			placed.toEnd++;
		}
		const last = result.at(-1);
		if (last && placed.fromStart <= last.fromEnd) {
			last.fromEnd = Math.max(last.fromEnd, placed.fromEnd);
			last.toEnd = Math.max(last.toEnd, placed.toEnd);
		} else result.push(placed);
	}
	return result;
}
