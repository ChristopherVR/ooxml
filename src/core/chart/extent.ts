// The smallest and largest of a chart's values without spreading them into one call: a spread
// passes one argument per value and throws once a series outgrows the engine's argument limit.

/**
 * The smallest and largest of `values`, exactly as `Math.min(...values)` and `Math.max(...values)`
 * give them (`Infinity` and `-Infinity` when there are none, `NaN` when any value is `NaN`).
 */
export function extentOf(values: Iterable<number>): { min: number; max: number } {
	let min = Infinity;
	let max = -Infinity;
	for (const value of values) {
		min = Math.min(min, value);
		max = Math.max(max, value);
	}
	return { min, max };
}
