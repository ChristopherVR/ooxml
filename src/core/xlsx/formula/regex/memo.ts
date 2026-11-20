import { type Program, REPEAT } from './compile';

/**
 * The visited set of the matcher: one bit per (memo point, `k`, position), one table per scope.
 * A state is marked when it is entered; meeting it again means it already failed (the search
 * stops at the first success, and the progress rule leaves no cycle that returns to the same
 * state), so it is pruned. That bounds a search by program size times input length.
 */

/** The bits all scopes of one search may use (8 MiB); scopes past it run without memo. */
const MAX_MEMO_BITS = 1 << 26;

let program: Program;
let length = 0;
let tables: (Uint32Array | null)[] = [];
let lo: number[] = [];
let hi: number[] = [];
const pool: Uint32Array[] = [];

/** Turns the memo on for a search of `program` over `n` code units; tables come zeroed. */
export function startMemo(p: Program, n: number): void {
	program = p;
	length = n;
	let left = MAX_MEMO_BITS;
	tables = p.strides.map((stride, scope) => {
		const need = stride * (n + 1);
		if (stride === 0 || need > left) return null;
		left -= need;
		const words = (need + 31) >>> 5;
		let bits = pool[scope];
		if (!bits || bits.length < words) pool[scope] = bits = new Uint32Array(words);
		return bits;
	});
	lo = p.strides.map(() => Infinity);
	hi = p.strides.map(() => -1);
}

/** Whether memo point `pc` (slot `extra`) was already entered at `pos`; marks it entered. */
export function seen(pc: number, pos: number, extra: number, regs: Int32Array): boolean {
	const base = program.memo[pc] as number;
	const scope = program.scope[pc] as number;
	const bits = tables[scope];
	if (base < 0 || !bits || pos > length) return false;
	// k: how many enclosing iterations (innermost first) started here and have not consumed.
	let k = 0;
	const marks = program.marks;
	for (let j = program.marksFrom[pc] as number, to = program.marksTo[pc] as number; j < to; j++) {
		if (regs[marks[j] as number] !== pos) break;
		k++;
	}
	const slots = program.op[pc] === REPEAT ? 2 : 1;
	const slot = pos * (program.strides[scope] as number) + base + k * slots + extra;
	const word = slot >>> 5;
	const bit = 1 << (slot & 31);
	const value = bits[word] as number;
	if (value & bit) return true;
	bits[word] = value | bit;
	if (pos < (lo[scope] as number)) lo[scope] = pos;
	if (pos > (hi[scope] as number)) hi[scope] = pos;
	return false;
}

/** Zeroes what a scope marked (a lookaround body is matched afresh at each position). */
export function clearMemo(scope: number): void {
	const bits = tables[scope];
	const from = lo[scope] as number;
	const to = hi[scope] as number;
	if (!bits || to < from) return;
	const stride = program.strides[scope] as number;
	bits.fill(0, (from * stride) >>> 5, (((to + 1) * stride - 1) >>> 5) + 1);
	lo[scope] = Infinity;
	hi[scope] = -1;
}

export function stopMemo(): void {
	for (let scope = 0; scope < tables.length; scope++) clearMemo(scope);
	tables = [];
}
