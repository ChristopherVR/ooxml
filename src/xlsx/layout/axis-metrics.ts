/**
 * Offsets along one grid axis (rows or columns) where nearly every line has the default size and
 * a sparse set of lines is custom-sized or hidden (size 0). Lookups are O(log n) in the number of
 * custom lines; nothing proportional to the 1,048,576 rows of a sheet is ever allocated.
 */
export class AxisMetrics {
	/** Indices of custom-sized lines, ascending. */
	private readonly indices: number[];
	/** Size of each custom line. */
	private readonly sizes: number[];
	/** `prefix[k]`: sum of `(size - defaultSize)` over custom lines before `indices[k]`. */
	private readonly prefix: number[];

	constructor(
		readonly defaultSize: number,
		/** Number of lines on the axis (last index + 1). */
		readonly count: number,
		overrides: Iterable<readonly [index: number, size: number]>,
	) {
		const sorted = [...overrides]
			.filter(([index, size]) => index >= 0 && index < count && size !== defaultSize)
			.sort((a, b) => a[0] - b[0]);
		this.indices = [];
		this.sizes = [];
		this.prefix = [];
		let delta = 0;
		for (const [index, size] of sorted) {
			if (this.indices.at(-1) === index) continue;
			this.indices.push(index);
			this.sizes.push(Math.max(0, size));
			this.prefix.push(delta);
			delta += Math.max(0, size) - defaultSize;
		}
		this.totalDelta = delta;
	}

	private readonly totalDelta: number;

	/** Number of custom entries strictly before `index` (binary search). */
	private rank(index: number): number {
		let lo = 0;
		let hi = this.indices.length;
		while (lo < hi) {
			const mid = (lo + hi) >> 1;
			if ((this.indices[mid] ?? 0) < index) lo = mid + 1;
			else hi = mid;
		}
		return lo;
	}

	size(index: number): number {
		if (index < 0 || index >= this.count) return 0;
		const k = this.rank(index);
		return this.indices[k] === index ? (this.sizes[k] ?? 0) : this.defaultSize;
	}

	/** Offset of the leading edge of `index` (`index` may equal `count` for the axis end). */
	start(index: number): number {
		const i = Math.max(0, Math.min(index, this.count));
		const k = this.rank(i);
		const before = k < this.indices.length ? (this.prefix[k] ?? 0) : this.totalDelta;
		return i * this.defaultSize + before;
	}

	/** Total extent of lines `0..count-1`. */
	get total(): number {
		return this.start(this.count);
	}

	/**
	 * The line containing `offset`: the last line whose start is at or before it and whose size is
	 * not zero. Offsets before the axis give 0, offsets past the end give the last visible line.
	 */
	at(offset: number): number {
		if (offset <= 0) return this.firstVisible(0);
		if (offset >= this.total) return this.lastVisible(this.count - 1);
		// Last custom entry starting at or before `offset`.
		let lo = 0;
		let hi = this.indices.length;
		while (lo < hi) {
			const mid = (lo + hi) >> 1;
			const idx = this.indices[mid] ?? 0;
			if (idx * this.defaultSize + (this.prefix[mid] ?? 0) <= offset) lo = mid + 1;
			else hi = mid;
		}
		const k = lo - 1;
		if (k < 0) return this.defaultSize > 0 ? Math.floor(offset / this.defaultSize) : 0;
		const idx = this.indices[k] ?? 0;
		const top = idx * this.defaultSize + (this.prefix[k] ?? 0);
		const size = this.sizes[k] ?? 0;
		if (offset < top + size) return idx;
		if (this.defaultSize <= 0) return this.firstVisible(idx + 1);
		const line = idx + 1 + Math.floor((offset - top - size) / this.defaultSize);
		return Math.min(line, this.count - 1);
	}

	/** The first line at or after `index` with a non-zero size (or `count - 1`). */
	firstVisible(index: number): number {
		for (let i = Math.max(0, index); i < this.count; i++) if (this.size(i) > 0) return i;
		return this.count - 1;
	}

	/** The last line at or before `index` with a non-zero size (or 0). */
	lastVisible(index: number): number {
		for (let i = Math.min(index, this.count - 1); i >= 0; i--) if (this.size(i) > 0) return i;
		return 0;
	}

	/** Largest custom-entry index, or -1. */
	get lastCustom(): number {
		return this.indices.at(-1) ?? -1;
	}
}
