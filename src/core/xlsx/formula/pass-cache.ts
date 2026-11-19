// What one evaluation pass reads once and reuses: range blocks, exact-match searches and the
// results of aggregates over ranges. Every entry describes cells whose formulas are already
// computed, so it stays true until a stored value changes under it (see `PassCache.clear`).
import { round15 } from './text-number';
import type { Matrix, Scalar, Value } from './values';

type ExactKey = number | string | boolean | null;

/**
 * The value `exactMatcher` treats as equal to `value`, for hashing: lower-cased text, and numbers
 * at the 15 significant digits `compareNumbers` compares (two numbers that round alike are always
 * within its closeness bound, so equal keys and equal numbers are the same thing). Errors and NaN
 * equal nothing, so they have no key.
 */
function exactKey(value: Scalar): ExactKey | undefined {
	if (typeof value === 'number') {
		if (Number.isNaN(value)) return undefined;
		return Number.isInteger(value) && Math.abs(value) < 1e15 ? value : round15(value);
	}
	if (typeof value === 'string') return value.toLowerCase();
	return typeof value === 'boolean' || value === null ? value : undefined;
}

/**
 * A linear exact search that remembers where it has been. It reads cells in the order a scan
 * would and stops at the same first match, so a cell a scan would never reach (a formula that
 * depends on the caller, say) is not read; the next search resumes where this one stopped.
 */
export class ExactSearch {
	private readonly first = new Map<ExactKey, number>();
	private visited = 0;
	/** Set while `find` reads cells: a formula it computes may search the same cells again. */
	busy = false;

	constructor(
		private readonly length: number,
		private readonly reverse: boolean,
	) {}

	find(lookup: Scalar, read: (index: number) => Scalar): number {
		const target = exactKey(lookup);
		const known = target === undefined ? undefined : this.first.get(target);
		if (known !== undefined) return known;
		this.busy = true;
		try {
			return this.extend(target, read);
		} finally {
			this.busy = false;
		}
	}

	private extend(target: ExactKey | undefined, read: (index: number) => Scalar): number {
		while (this.visited < this.length) {
			const at = this.reverse ? this.length - 1 - this.visited : this.visited;
			// Counted after the read: one that unwinds to compute a formula first is repeated.
			const key = exactKey(read(at));
			this.visited++;
			if (key === undefined) continue;
			if (!this.first.has(key)) this.first.set(key, at);
			if (key === target) return at;
		}
		return -1;
	}
}

/** What the cache needs to know about the engine that owns it. */
interface PassState {
	/** Whether an evaluation pass is running whose reads can be kept. */
	active(): boolean;
	/** How many reads so far hit a formula still being evaluated. */
	cycleReads(): number;
}

export class PassCache {
	readonly blocks = new Map<string, Matrix>();
	/** By vector; `null` until its second search, so a lone lookup stays a plain scan. */
	readonly searches = new Map<string, ExactSearch | null>();
	readonly results = new Map<string, Value>();
	/** Changes on every `clear`, so a read that spans one is not stored afterwards. */
	private generation = 0;

	constructor(private readonly state: PassState) {}

	/** Forgets everything: the pass is over or a value already read has changed. */
	clear(): void {
		this.generation++;
		this.blocks.clear();
		this.searches.clear();
		this.results.clear();
	}

	/** `read()`, kept in `entries` unless it saw a formula in progress or a `clear`. */
	remember<T>(entries: Map<string, T>, key: string, read: () => T): T {
		if (!this.state.active()) return read();
		const known = entries.get(key);
		if (known !== undefined) return known;
		const { generation } = this;
		const cycleReads = this.state.cycleReads();
		const value = read();
		if (this.generation === generation && this.state.cycleReads() === cycleReads)
			entries.set(key, value);
		return value;
	}

	/** An exact search through the index of `vector`; undefined when the caller should scan. */
	findExact(
		vector: string,
		length: number,
		read: (index: number) => Scalar,
		lookup: Scalar,
		reverse: boolean,
	): number | undefined {
		if (!this.state.active()) return undefined;
		const key = `${reverse ? 'r' : 'f'}${vector}`;
		let search = this.searches.get(key);
		if (search === undefined) {
			this.searches.set(key, null);
			return undefined;
		}
		if (!search) this.searches.set(key, (search = new ExactSearch(length, reverse)));
		// A lookup computed while this index reads its cells scans on its own, as it would alone.
		else if (search.busy) return undefined;
		const cycleReads = this.state.cycleReads();
		try {
			return search.find(lookup, read);
		} finally {
			if (this.state.cycleReads() !== cycleReads) this.searches.delete(key);
		}
	}
}
