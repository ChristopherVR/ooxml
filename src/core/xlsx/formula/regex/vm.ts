import { chargeWork, ERR, fail } from '../values';
import { type CharMatcher, codePointAt, codePointBefore, width } from './charset';
import {
	ASSERT,
	BACKREF,
	CHAR,
	CINC,
	CINIT,
	CLOOP,
	CPROGRESS,
	JMP,
	LOOK,
	MARK,
	MATCH,
	PROGRESS,
	type Program,
	REPEAT,
	RESET,
	SAVE,
	SPLIT,
} from './compile';
import { assertion, backref } from './checks';
import { clearMemo, seen as memoSeen, startMemo, stopMemo } from './memo';

/**
 * Work units charged per matcher step (an instruction, a character read, or a backtrack). A unit
 * is about 4 ns of the recalculation budget and a step measured 6 to 16 ns, so a pattern that
 * spends the whole budget stops within one to two seconds, as other heavy formulas do.
 */
const STEP_WORK = 4;
const CHUNK = 4096;
let chunk = CHUNK;
/** The backtrack stack's limit in 32-bit words (16 MiB); past it the match is `#NUM!`. */
const MAX_STACK = 1 << 22;
/**
 * Searches start without the memo, which costs a bit per step; it is turned on once a search has
 * taken more steps than an ordinary pattern needs, and from then on no state is explored twice.
 */
const MEMO_AFTER = (scanned: number): number => 4096 + 8 * scanned;
let memoAfter = MEMO_AFTER;

// Backtrack entries, four words each: kind, then operands.
const CHOICE = 0; // pc, position
const UNDO = 1; // register, old value
const GREEDY = 2; // REPEAT pc, end, shortest end
const LAZY = 3; // REPEAT pc, end, count

let p: Program;
let input = '';
let n = 0;
let regs = new Int32Array(0);
let stack = new Int32Array(1024);
let sp = 0;
let steps = 0;
let searchSteps = 0;
let memoAt = 0;
let memoOn = false;
/** Where the leading REPEAT's scan stopped in the current attempt, or -1. */
let leadEnd = -1;

/** For tests: when the memo starts (by steps taken) and how often the budget is charged. */
export function tuneMemo(after = MEMO_AFTER, every = CHUNK): void {
	memoAfter = after;
	chunk = every;
}

function flush(): void {
	const units = steps * STEP_WORK;
	searchSteps += steps;
	steps = 0;
	if (!memoOn && p.memoize && searchSteps >= memoAt) {
		startMemo(p, n);
		memoOn = true;
	}
	chargeWork(units);
}

const seen = (pc: number, pos: number, extra: number): boolean =>
	memoOn && memoSeen(pc, pos, extra, regs);

function push(kind: number, x: number, y: number, z: number): void {
	if (sp + 4 > stack.length) {
		if (stack.length >= MAX_STACK) fail(ERR.NUM);
		const grown = new Int32Array(stack.length * 2);
		grown.set(stack);
		stack = grown;
	}
	stack[sp] = kind;
	stack[sp + 1] = x;
	stack[sp + 2] = y;
	stack[sp + 3] = z;
	sp += 4;
}

/** Writes a register so that backtracking restores it. */
function write(register: number, value: number): void {
	push(UNDO, register, regs[register] as number, 0);
	regs[register] = value;
}

/** The position after one code point matching `m` from `q` (backward when `back`), or -1. */
function step(m: CharMatcher, q: number, back: boolean): number {
	steps++;
	if (back) {
		if (q <= 0) return -1;
		const cp = codePointBefore(input, q);
		return m.test(cp) ? q - width(cp) : -1;
	}
	if (q >= n) return -1;
	const cp = codePointAt(input, q);
	return m.test(cp) ? q + width(cp) : -1;
}

/** `q` moved one code point back toward where the repeat started. */
const unstep = (q: number, back: boolean): number =>
	back ? q + width(codePointAt(input, q)) : q - width(codePointBefore(input, q));

/** Whether the CHAR after REPEAT `pc` (if any) can match at `q`. */
function continues(pc: number, q: number, back: boolean): boolean {
	const next = ((p.d[pc] as number) >> 3) - 1;
	if (next < 0) return true;
	const m = p.matchers[next] as CharMatcher;
	if (back) return q > 0 && m.test(codePointBefore(input, q));
	return q < n && m.test(codePointAt(input, q));
}

/** A greedy REPEAT tries ends from `q` back to `shortest`; returns the end to continue at. */
function greedyEnd(pc: number, q: number, shortest: number, back: boolean): number {
	// An unbounded repeat marked every end while scanning; a bounded one marks each end it tries.
	const unbounded = (p.c[pc] as number) < 0;
	// Inside the run every end is followed by a character of the run, which the literal is not.
	const longestOnly = ((p.d[pc] as number) & 4) !== 0;
	for (;;) {
		if (continues(pc, q, back) && (unbounded || !seen(pc, q, 1))) {
			if (q !== shortest && !longestOnly) push(GREEDY, pc, q, shortest);
			return q;
		}
		if (longestOnly) return -1;
		if (q === shortest) return -1;
		q = unstep(q, back);
		steps++;
	}
}

/** A lazy REPEAT continues at `q` first, keeping the option to take one more. */
function lazyEnd(pc: number, q: number, count: number, back: boolean): number {
	const max = p.c[pc] as number;
	const m = p.matchers[p.a[pc] as number] as CharMatcher;
	for (;;) {
		// An unbounded repeat is in the same state at `q` whatever the count.
		if (max < 0 && seen(pc, q, 1)) return -1;
		if (continues(pc, q, back) && (max < 0 || !seen(pc, q, 1))) {
			if (max < 0 || count < max) push(LAZY, pc, q, count);
			return q;
		}
		if (max >= 0 && count >= max) return -1;
		if ((q = step(m, q, back)) < 0) return -1;
		count++;
	}
}

function repeat(pc: number, start: number): number {
	if (seen(pc, start, 0)) return -1;
	const m = p.matchers[p.a[pc] as number] as CharMatcher;
	const min = p.b[pc] as number;
	const max = (p.c[pc] as number) < 0 ? Infinity : (p.c[pc] as number);
	const back = ((p.d[pc] as number) & 2) !== 0;
	let q = start;
	for (let i = 0; i < min; i++) if ((q = step(m, q, back)) < 0) return -1;
	if (!((p.d[pc] as number) & 1)) return lazyEnd(pc, q, min, back);
	const shortest = q;
	let count = min;
	if (memoOn && max === Infinity) {
		for (;;) {
			if (seen(pc, q, 1)) {
				// The ends from here on were tried by an earlier visit.
				if (pc === p.lead) leadEnd = q;
				if (q === shortest) return -1;
				q = unstep(q, back);
				break;
			}
			const next = step(m, q, back);
			if (next < 0) break;
			q = next;
		}
	} else if (!back) {
		const ascii = m.ascii;
		while (count < max && q < n) {
			const unit = input.charCodeAt(q);
			if (unit < 128) {
				if (ascii[unit] === 0) break;
				q++;
			} else {
				const cp = unit >= 0xd800 && unit < 0xdc00 ? codePointAt(input, q) : unit;
				if (!m.test(cp)) break;
				q += width(cp);
			}
			count++;
		}
		steps += count - min;
	} else {
		for (let next: number; count < max && (next = step(m, q, back)) >= 0; count++) q = next;
	}
	if (pc === p.lead) leadEnd = q;
	return greedyEnd(pc, q, shortest, back);
}

function resume(kind: number, pc: number, q: number, extra: number): number {
	const back = ((p.d[pc] as number) & 2) !== 0;
	if (kind === GREEDY) return greedyEnd(pc, unstep(q, back), extra, back);
	const next = step(p.matchers[p.a[pc] as number] as CharMatcher, q, back);
	return next < 0 ? -1 : lazyEnd(pc, next, extra + 1, back);
}

/** Runs a lookaround body on its own; a positive one keeps its captures. */
function look(pc: number, pos: number): boolean {
	const flags = p.b[pc] as number;
	const saved = regs.slice();
	let end: number;
	try {
		end = run(p.a[pc] as number, pos);
	} finally {
		if (memoOn) clearMemo(p.c[pc] as number);
	}
	if (flags & 2) {
		if (end >= 0) regs.set(saved);
		return end < 0;
	}
	if (end < 0) return false;
	for (let r = 0; r < regs.length; r++)
		if (regs[r] !== saved[r]) push(UNDO, r, saved[r] as number, 0);
	return true;
}

/** Matches from instruction `pc` at `pos`: the end position, or -1. */
function run(pc: number, pos: number): number {
	const base = sp;
	const { op, a, b, c, d, matchers } = p;
	for (;;) {
		if (++steps >= chunk) flush();
		let ok = true;
		switch (op[pc]) {
			case CHAR: {
				const m = matchers[a[pc] as number] as CharMatcher;
				if (b[pc] === 0) {
					const cp = pos < n ? codePointAt(input, pos) : -1;
					if ((ok = cp >= 0 && m.test(cp))) pos += width(cp);
				} else {
					const cp = pos > 0 ? codePointBefore(input, pos) : -1;
					if ((ok = cp >= 0 && m.test(cp))) pos -= width(cp);
				}
				pc++;
				break;
			}
			case SPLIT:
				if ((ok = !seen(pc, pos, 0))) {
					push(CHOICE, b[pc] as number, pos, 0);
					pc = a[pc] as number;
				}
				break;
			case JMP:
				pc = a[pc] as number;
				break;
			case SAVE:
			case MARK:
				write(a[pc++] as number, pos);
				break;
			case RESET:
				for (let r = a[pc] as number; r < (b[pc] as number); r++) if (regs[r] !== -1) write(r, -1);
				pc++;
				break;
			case ASSERT:
				ok = assertion(input, a[pc] as number, matchers[b[pc] as number], pos);
				pc++;
				break;
			case BACKREF: {
				const end = backref(
					input,
					regs,
					p.refs[a[pc] as number] ?? [],
					b[pc] === 1,
					c[pc] === 1,
					pos,
				);
				steps += Math.abs(end - pos);
				if ((ok = end >= 0)) pos = end;
				pc++;
				break;
			}
			case LOOK:
				ok = look(pc++, pos);
				break;
			case MATCH:
				sp = base;
				return pos;
			case PROGRESS:
				ok = regs[a[pc++] as number] !== pos;
				break;
			case REPEAT: {
				const end = repeat(pc++, pos);
				if ((ok = end >= 0)) pos = end;
				break;
			}
			case CINIT:
				write(a[pc++] as number, 0);
				break;
			case CLOOP: {
				const count = regs[a[pc] as number] as number;
				const max = c[pc] as number;
				const exit = (d[pc] as number) >> 1;
				if (count < (b[pc] as number)) pc++;
				else if (max >= 0 && count >= max) pc = exit;
				else if ((d[pc] as number) & 1) {
					push(CHOICE, exit, pos, 0);
					pc++;
				} else {
					push(CHOICE, pc + 1, pos, 0);
					pc = exit;
				}
				break;
			}
			case CPROGRESS:
				ok = (regs[b[pc] as number] as number) < (c[pc] as number) || regs[a[pc] as number] !== pos;
				pc++;
				break;
			case CINC:
				write(a[pc] as number, (regs[a[pc++] as number] as number) + 1);
				break;
		}
		if (ok) continue;
		// Backtrack to the latest choice, undoing register writes on the way.
		for (;;) {
			if (sp <= base) return -1;
			if (++steps >= chunk) flush();
			sp -= 4;
			const kind = stack[sp] as number;
			const x = stack[sp + 1] as number;
			const y = stack[sp + 2] as number;
			if (kind === UNDO) {
				regs[x] = y;
				continue;
			}
			if (kind === CHOICE) {
				pc = x;
				pos = y;
				break;
			}
			const end = resume(kind, x, y, stack[sp + 3] as number);
			if (end >= 0) {
				pc = x + 1;
				pos = end;
				break;
			}
		}
	}
}

/** The first position at or after `start` where a match can begin, by the program's first atoms. */
function candidate(start: number): number {
	const table = p.firstAscii;
	const firsts = p.first;
	if (!table || !firsts) return start;
	let at = start;
	for (; at < n; at++) {
		const unit = input.charCodeAt(at);
		if (unit < 128) {
			if (table[unit]) break;
			continue;
		}
		const cp = codePointAt(input, at);
		if (firsts.some((m) => m.test(cp))) break;
		if (cp > 0xffff) at++;
	}
	steps += (at - start) >> 2;
	return at;
}

/** Start and end of a match, then of each group (-1 when the group did not take part). */
export type Match = readonly number[];

/**
 * The leftmost match of `program` in `text` starting at or after `from` (a code point boundary):
 * its {@link Match}, or null.
 * Every step is charged to the recalculation's work budget, so it is `#NUM!` rather than a hang.
 */
export function search(program: Program, text: string, from: number): Match | null {
	if (p !== program || regs.length !== program.registers) regs = new Int32Array(program.registers);
	p = program;
	input = text;
	n = text.length;
	regs.fill(-1);
	sp = 0;
	searchSteps = 0;
	memoAt = memoAfter(n - from);
	memoOn = false;
	try {
		for (let start = candidate(from); start <= n; start = candidate(start)) {
			if (program.first && start >= n) break;
			leadEnd = -1;
			const end = run(0, start);
			if (end >= 0) {
				regs[0] = start;
				regs[1] = end;
				const match: number[] = [];
				for (let r = 0, to = 2 * (program.groupCount + 1); r < to; r++)
					match.push(regs[r] as number);
				return match;
			}
			if (program.anchored) break;
			start = leadEnd > start ? leadEnd : start + (start < n ? width(codePointAt(text, start)) : 1);
		}
		return null;
	} finally {
		sp = 0;
		if (memoOn) stopMemo();
		memoOn = false;
		const units = steps * STEP_WORK;
		steps = 0;
		chargeWork(units);
	}
}
