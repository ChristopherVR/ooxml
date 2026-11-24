/** Tag trees (ITU-T T.800 B.10.2): the code-block inclusion and zero bit-plane counts of a precinct. */
export class TagTree {
	private readonly levels: { width: number; values: Int32Array; lows: Int32Array }[] = [];

	constructor(width: number, height: number) {
		for (;;) {
			const size = Math.max(1, width * height);
			this.levels.push({
				width,
				values: new Int32Array(size).fill(0x7fffffff),
				lows: new Int32Array(size),
			});
			if (size === 1) break;
			width = Math.ceil(width / 2);
			height = Math.ceil(height / 2);
		}
	}

	/**
	 * Reads bits until the value of leaf (x, y) is known or reaches `threshold`; returns the value
	 * when it is below the threshold, otherwise undefined.
	 */
	decode(x: number, y: number, threshold: number, bit: () => number): number | undefined {
		let low = 0;
		for (let level = this.levels.length - 1; level >= 0; level--) {
			const { width, values, lows } = this.levels[level]!;
			const index = (y >> level) * width + (x >> level);
			if (low > lows[index]!) lows[index] = low;
			else low = lows[index]!;
			while (low < threshold && low < values[index]!) {
				if (bit()) values[index] = low;
				else low++;
			}
			lows[index] = low;
		}
		const value = this.levels[0]!.values[y * this.levels[0]!.width + x]!;
		return value < threshold ? value : undefined;
	}
}
