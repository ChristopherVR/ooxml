// Sequence difference by Myers' O(ND) algorithm ("An O(ND) Difference Algorithm and Its
// Variations", 1986), after trimming the common prefix and suffix.

/** One step of an edit script: `equal` pairs `a[aIndex]` with `b[bIndex]`. */
export type DiffOp =
	| { kind: 'equal'; aIndex: number; bIndex: number }
	| { kind: 'delete'; aIndex: number }
	| { kind: 'insert'; bIndex: number };

/** Edit distances above this fall back to replacing the whole middle, bounding time and memory. */
const MAX_EDIT_DISTANCE = 4000;

function middleScript(a: readonly string[], b: readonly string[], offset: number): DiffOp[] {
	const n = a.length;
	const m = b.length;
	const max = Math.min(n + m, MAX_EDIT_DISTANCE);
	const size = 2 * max + 1;
	const v = new Int32Array(size + 2);
	const trace: Int32Array[] = [];
	let found = -1;
	for (let d = 0; d <= max && found < 0; d++) {
		trace.push(v.slice());
		for (let k = -d; k <= d; k += 2) {
			const index = k + max;
			let x =
				k === -d || (k !== d && v[index - 1]! < v[index + 1]!) ? v[index + 1]! : v[index - 1]! + 1;
			let y = x - k;
			while (x < n && y < m && a[x] === b[y]) {
				x++;
				y++;
			}
			v[index] = x;
			if (x >= n && y >= m) {
				found = d;
				break;
			}
		}
	}
	if (found < 0) {
		const ops: DiffOp[] = [];
		for (let i = 0; i < n; i++) ops.push({ kind: 'delete', aIndex: offset + i });
		for (let j = 0; j < m; j++) ops.push({ kind: 'insert', bIndex: offset + j });
		return ops;
	}
	const reversed: DiffOp[] = [];
	let x = n;
	let y = m;
	for (let d = found; d > 0; d--) {
		const previous = trace[d]!;
		const k = x - y;
		const down = k === -d || (k !== d && previous[k - 1 + max]! < previous[k + 1 + max]!);
		const prevK = down ? k + 1 : k - 1;
		const prevX = previous[prevK + max]!;
		const prevY = prevX - prevK;
		while (x > prevX && y > prevY) {
			x--;
			y--;
			reversed.push({ kind: 'equal', aIndex: offset + x, bIndex: offset + y });
		}
		if (down) reversed.push({ kind: 'insert', bIndex: offset + --y });
		else reversed.push({ kind: 'delete', aIndex: offset + --x });
	}
	while (x > 0 && y > 0) {
		x--;
		y--;
		reversed.push({ kind: 'equal', aIndex: offset + x, bIndex: offset + y });
	}
	return reversed.reverse();
}

/**
 * The shortest edit script turning `a` into `b`, comparing items by their keys. Deletions come
 * before insertions where both are possible, which is how Word orders a replaced span.
 */
export function diffKeys(a: readonly string[], b: readonly string[]): DiffOp[] {
	let start = 0;
	while (start < a.length && start < b.length && a[start] === b[start]) start++;
	let endA = a.length;
	let endB = b.length;
	while (endA > start && endB > start && a[endA - 1] === b[endB - 1]) {
		endA--;
		endB--;
	}
	const ops: DiffOp[] = [];
	for (let i = 0; i < start; i++) ops.push({ kind: 'equal', aIndex: i, bIndex: i });
	if (endA > start || endB > start) {
		const middle = middleScript(a.slice(start, endA), b.slice(start, endB), start);
		ops.push(...middle);
	}
	for (let i = 0; endA + i < a.length; i++)
		ops.push({ kind: 'equal', aIndex: endA + i, bIndex: endB + i });
	return orderReplacements(ops);
}

/** Within each run of changes, every deletion before every insertion. */
export function orderReplacements(ops: DiffOp[]): DiffOp[] {
	const result: DiffOp[] = [];
	let deletes: DiffOp[] = [];
	let inserts: DiffOp[] = [];
	const flush = () => {
		result.push(...deletes, ...inserts);
		deletes = [];
		inserts = [];
	};
	for (const op of ops) {
		if (op.kind === 'delete') deletes.push(op);
		else if (op.kind === 'insert') inserts.push(op);
		else {
			flush();
			result.push(op);
		}
	}
	flush();
	return result;
}
