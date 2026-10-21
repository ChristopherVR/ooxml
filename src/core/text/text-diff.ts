// The changes that turn one string into another: a shortest edit script from Myers' O(ND) diff
// ("An O(ND) Difference Algorithm and Its Variations", 1986), over UTF-16 code units. No change
// starts or ends inside a surrogate pair.

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
 * The insertions and deletions of a shortest edit script from `a` to `b`, as `[x, y, inserts]`:
 * the edit at `a[x]` (a deletion) or `b[y]` (an insertion). Undefined past `maxEdits` edits.
 */
function shortestEdit(
	a: string,
	b: string,
	maxEdits: number,
): [number, number, boolean][] | undefined {
	const n = a.length;
	const m = b.length;
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
			while (x < n && y < m && a.charCodeAt(x) === b.charCodeAt(y)) {
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

/**
 * The changes from `from` to `to`, in order and separated by unchanged text. When the strings
 * differ by more than `maxEdits` inserted or deleted code units, everything between their common
 * start and end is reported as one change.
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
	const edits = a.length && b.length ? shortestEdit(a, b, maxEdits) : undefined;
	const changes: TextChange[] = [];
	if (!edits) changes.push({ fromStart: 0, fromEnd: a.length, toStart: 0, toEnd: b.length });
	for (const [x, y, inserts] of edits ?? []) {
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
