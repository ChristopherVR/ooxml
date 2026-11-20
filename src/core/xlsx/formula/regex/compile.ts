import type { CharMatcher } from './charset';
import { anchored, firstAtoms, nullable } from './analysis';
import type { Node, Parsed } from './parse';

// Instructions. Operands are `a` to `d`; a SPLIT prefers `a` and backtracks to `b`.
export const CHAR = 0; // a matcher, b direction (1 backward)
export const SPLIT = 1;
export const JMP = 2; // a target
export const SAVE = 3; // a register
export const RESET = 4; // clears registers [a, b)
export const ASSERT = 5; // a Assertion, b word matcher
export const BACKREF = 6; // a group list, b direction, c ignore case
export const LOOK = 7; // a body, b flags (1 behind, 2 negative), c scope
export const MATCH = 8;
export const MARK = 9; // a register := position
export const PROGRESS = 10; // fails when register a still holds the position
// REPEAT: a matcher, b min, c max (-1 unbounded), d flags: 1 greedy, 2 backward, 4 when the CHAR
// that follows (after group ends) is a literal the repeat cannot match (so only the longest run can continue), and
// from bit 3 one more than that CHAR's matcher, which lets it skip ends that cannot continue.
export const REPEAT = 11;
export const CINIT = 12; // a counter := 0
export const CLOOP = 13; // a counter, b min, c max (-1 unbounded), d exit * 2 + greedy
export const CPROGRESS = 14; // like PROGRESS on register a, once counter b has reached c
export const CINC = 15; // a counter += 1

/**
 * A compiled pattern. Each SPLIT and REPEAT is a memo point: the matcher records which
 * (point, position) states it has tried, so no state is explored twice. Iterations that may match
 * empty fail when they do (as in JavaScript), which needs their start position, so a state also
 * records how many of the enclosing such iterations have not consumed yet (`k`).
 */
export interface Program {
	readonly op: Int32Array;
	readonly a: Int32Array;
	readonly b: Int32Array;
	readonly c: Int32Array;
	readonly d: Int32Array;
	/** First memo slot of a point within its scope's row, or -1. */
	readonly memo: Int32Array;
	/** 0 for the pattern, then one scope per lookaround body (each is matched on its own). */
	readonly scope: Int32Array;
	/** The mark registers of a point's enclosing iterations, innermost first: `marks[from, to)`. */
	readonly marksFrom: Int32Array;
	readonly marksTo: Int32Array;
	readonly marks: Int32Array;
	/** Memo slots per input position, per scope. */
	readonly strides: readonly number[];
	/** False with backreferences or counters, whose state a (point, position) pair does not hold. */
	readonly memoize: boolean;
	readonly matchers: readonly CharMatcher[];
	readonly refs: readonly (readonly number[])[];
	readonly registers: number;
	readonly groupCount: number;
	readonly names: ReadonlyMap<string, readonly number[]>;
	/** The atoms a match must start with, or null when it may start anywhere. */
	readonly first: readonly CharMatcher[] | null;
	/** For ASCII code units, whether a match can start with one (null with `first`). */
	readonly firstAscii: Uint8Array | null;
	/** Whether a match can only start at position 0. */
	readonly anchored: boolean;
	/**
	 * A greedy unbounded REPEAT the pattern starts with (after group starts), or -1. When a match
	 * from one start fails, no start inside the run it consumed can succeed: each would try a subset
	 * of the same ends, whose continuations do not depend on where the match started.
	 */
	readonly lead: number;
}

/** Past this many instructions, counted repeats are not unrolled but counted (without memo). */
const MAX_UNROLLED = 20_000;
/** Bigger programs than this are refused. */
const MAX_PROGRAM = 200_000;
class TooLarge extends Error {}

interface Context {
	readonly back: boolean;
	readonly marks: number[];
}

class Builder {
	readonly op: number[] = [];
	readonly a: number[] = [];
	readonly b: number[] = [];
	readonly c: number[] = [];
	readonly d: number[] = [];
	readonly memo: number[] = [];
	readonly scope: number[] = [];
	readonly marksFrom: number[] = [];
	readonly marksTo: number[] = [];
	readonly marks: number[] = [];
	readonly strides = [0];
	readonly matchers: CharMatcher[] = [];
	readonly refs: number[][] = [];
	readonly looks: { pc: number; node: Node; back: boolean; scope: number }[] = [];
	private readonly index = new Map<CharMatcher, number>();
	/** The scope being emitted. */
	current = 0;
	registers: number;
	hasRefs = false;

	constructor(
		groupCount: number,
		private readonly counters: boolean,
	) {
		this.registers = 2 * (groupCount + 1);
	}

	get here(): number {
		return this.op.length;
	}

	emit(op: number, a = 0, b = 0, c = 0, d = 0): number {
		if (this.op.length >= (this.counters ? MAX_PROGRAM : MAX_UNROLLED)) throw new TooLarge();
		this.op.push(op);
		this.a.push(a);
		this.b.push(b);
		this.c.push(c);
		this.d.push(d);
		this.memo.push(-1);
		this.scope.push(this.current);
		this.marksFrom.push(this.marks.length);
		this.marksTo.push(this.marks.length);
		return this.op.length - 1;
	}

	/** Makes `pc` a memo point with `slots` slots for each value of `k`. */
	point(pc: number, ctx: Context, slots: number): void {
		this.memo[pc] = this.strides[this.current] ?? 0;
		this.strides[this.current] = (this.memo[pc] ?? 0) + slots * (ctx.marks.length + 1);
		this.marksFrom[pc] = this.marks.length;
		for (let i = ctx.marks.length - 1; i >= 0; i--) this.marks.push(ctx.marks[i] ?? 0);
		this.marksTo[pc] = this.marks.length;
	}

	matcher(m: CharMatcher): number {
		let i = this.index.get(m);
		if (i === undefined) this.index.set(m, (i = this.matchers.push(m) - 1));
		return i;
	}

	branch(split: number, body: number, exit: number, greedy: boolean): void {
		this.a[split] = greedy ? body : exit;
		this.b[split] = greedy ? exit : body;
	}

	gen(node: Node, ctx: Context): void {
		switch (node.kind) {
			case 'char':
				this.emit(CHAR, this.matcher(node.matcher), ctx.back ? 1 : 0);
				return;
			case 'seq':
				for (const item of ctx.back ? [...node.items].reverse() : node.items) this.gen(item, ctx);
				return;
			case 'alt': {
				const jumps: number[] = [];
				const last = node.items.length - 1;
				for (const [i, item] of node.items.entries()) {
					const split = i < last ? this.emit(SPLIT, this.here + 1) : -1;
					if (split >= 0) this.point(split, ctx, 1);
					this.gen(item, ctx);
					if (split < 0) continue;
					jumps.push(this.emit(JMP));
					this.b[split] = this.here;
				}
				for (const jump of jumps) this.a[jump] = this.here;
				return;
			}
			case 'group': {
				const [open, close] = ctx.back ? [1, 0] : [0, 1];
				this.emit(SAVE, 2 * node.index + open);
				this.gen(node.body, ctx);
				this.emit(SAVE, 2 * node.index + close);
				return;
			}
			case 'assert':
				this.emit(ASSERT, node.what, node.word ? this.matcher(node.word) : -1);
				return;
			case 'backref':
				this.hasRefs = true;
				this.emit(
					BACKREF,
					this.refs.push(node.groups) - 1,
					ctx.back ? 1 : 0,
					node.ignoreCase ? 1 : 0,
				);
				return;
			case 'look': {
				const scope = this.strides.push(0) - 1;
				const pc = this.emit(LOOK, -1, (node.behind ? 1 : 0) | (node.negate ? 2 : 0), scope);
				this.looks.push({ pc, node: node.body, back: node.behind, scope });
				return;
			}
			case 'repeat':
				this.repeat(node, ctx);
		}
	}

	private repeat(node: Extract<Node, { kind: 'repeat' }>, ctx: Context): void {
		const { body, min, max, greedy } = node;
		if (max === 0) return;
		if (body.kind === 'char') {
			const pc = this.emit(
				REPEAT,
				this.matcher(body.matcher),
				min,
				max === Infinity ? -1 : max,
				(greedy ? 1 : 0) | (ctx.back ? 2 : 0),
			);
			this.point(pc, ctx, 2);
			return;
		}
		const [first, end] = node.groups;
		const empty = nullable(body);
		const reset = () => {
			if (end > first) this.emit(RESET, 2 * first, 2 * end);
		};
		/** One iteration that must consume (when its body could match empty). */
		const optional = (onEmpty: (mark: number) => void) => {
			reset();
			const mark = empty ? this.registers++ : -1;
			if (empty) {
				this.emit(MARK, mark);
				ctx.marks.push(mark);
			}
			this.gen(body, ctx);
			if (empty) {
				ctx.marks.pop();
				onEmpty(mark);
			}
		};
		if (this.counters && !(min === 0 && (max === 1 || max === Infinity))) {
			const counter = this.registers++;
			this.emit(CINIT, counter);
			const loop = this.emit(CLOOP, counter, min, max === Infinity ? -1 : max);
			optional((mark) => this.emit(CPROGRESS, mark, counter, min));
			this.emit(CINC, counter);
			this.emit(JMP, loop);
			this.d[loop] = this.here * 2 + (greedy ? 1 : 0);
			return;
		}
		for (let i = 0; i < min; i++) {
			reset();
			this.gen(body, ctx);
		}
		const splits: number[] = [];
		for (let i = min; i < (max === Infinity ? min + 1 : max); i++) {
			const split = this.emit(SPLIT);
			this.point(split, ctx, 1);
			splits.push(split);
			optional((mark) => this.emit(PROGRESS, mark));
			if (max === Infinity) this.emit(JMP, split);
		}
		for (const split of splits) this.branch(split, split + 1, this.here, greedy);
	}
}

function build(parsed: Parsed, counters: boolean): Program {
	const builder = new Builder(parsed.groupCount, counters);
	builder.gen(parsed.root, { back: false, marks: [] });
	builder.emit(MATCH);
	for (let i = 0; i < builder.looks.length; i++) {
		const look = builder.looks[i];
		if (!look) continue;
		builder.current = look.scope;
		builder.a[look.pc] = builder.here;
		builder.gen(look.node, { back: look.back, marks: [] });
		builder.emit(MATCH);
	}
	for (let pc = 0; pc < builder.here; pc++) {
		let next = pc + 1;
		while (builder.op[next] === SAVE) next++;
		const back = ((builder.d[pc] ?? 0) & 2) >> 1;
		if (builder.op[pc] !== REPEAT || builder.op[next] !== CHAR || builder.b[next] !== back)
			continue;
		const follower = builder.matchers[builder.a[next] ?? 0];
		const literal = follower?.literal ?? -1;
		const stops = literal >= 0 && !builder.matchers[builder.a[pc] ?? 0]?.test(literal);
		builder.d[pc] = (builder.d[pc] ?? 0) | (stops ? 4 : 0) | (((builder.a[next] ?? 0) + 1) << 3);
	}
	const starts = firstAtoms(parsed.root);
	const found = starts.set && !starts.empty ? [...new Set(starts.set)] : null;
	const unique = found && found.length <= 16 ? found : null;
	const firstAscii = unique
		? Uint8Array.from({ length: 128 }, (_, c) => (unique.some((m) => m.test(c)) ? 1 : 0))
		: null;
	const ints = (values: number[]) => Int32Array.from(values);
	let lead = 0;
	while (builder.op[lead] === SAVE) lead++;
	const leads =
		builder.op[lead] === REPEAT && builder.c[lead] === -1 && ((builder.d[lead] ?? 0) & 3) === 1;
	return {
		op: ints(builder.op),
		a: ints(builder.a),
		b: ints(builder.b),
		c: ints(builder.c),
		d: ints(builder.d),
		memo: ints(builder.memo),
		scope: ints(builder.scope),
		marksFrom: ints(builder.marksFrom),
		marksTo: ints(builder.marksTo),
		marks: ints(builder.marks),
		strides: builder.strides,
		memoize: !counters && !builder.hasRefs,
		matchers: builder.matchers,
		refs: builder.refs,
		registers: builder.registers,
		groupCount: parsed.groupCount,
		names: parsed.names,
		first: unique,
		firstAscii,
		anchored: anchored(parsed.root),
		lead: leads && !builder.hasRefs ? lead : -1,
	};
}

/**
 * Compiles a parsed pattern, unrolling counted repeats unless that makes the program too big (or
 * `unroll` is false), in which case they run on counters without memo.
 */
export function compile(parsed: Parsed, unroll = true): Program {
	try {
		if (unroll) return build(parsed, false);
	} catch (error) {
		if (!(error instanceof TooLarge)) throw error;
	}
	return build(parsed, true);
}
