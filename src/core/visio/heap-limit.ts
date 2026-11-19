/**
 * The default memory budget of a Visio read or edit, from the heap the host's runtime has.
 * Measured in Node 24 (README, Bounded processing).
 */
const MiB = 1024 * 1024;
/** Peak heap of a read or an edit for each byte its budget is charged, at most. */
export const PEAK_HEAP_PER_BUDGET_BYTE = 1.6;
/**
 * Heap left out of the budget: V8's young generation (192 MiB of Node's `heap_size_limit`, which
 * the old generation where parsed XML lives cannot use) and the host's own objects.
 */
export const HOST_HEAP_MARGIN = 256 * MiB;
/** The heap limit assumed where the runtime reports none (Firefox, Safari, workers). */
export const ASSUMED_HEAP_LIMIT = 1024 * MiB;
const MIN_BUDGET = 8 * MiB;

type Runtime = {
	process?: {
		versions?: { bun?: string };
		getBuiltinModule?: (id: string) => unknown;
	};
	performance?: { memory?: { jsHeapSizeLimit?: number } };
};

/**
 * The V8 heap limit of this runtime: Node and Deno report it (`v8.getHeapStatistics()`), Chromium
 * pages too (`performance.memory`). Bun's JavaScriptCore has no such limit, and its report is
 * not one.
 */
export function runtimeHeapLimit(runtime: Runtime = globalThis as Runtime): number | undefined {
	const valid = (value: unknown) =>
		typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : undefined;
	if (!runtime.process?.versions?.bun) {
		try {
			const v8 = runtime.process?.getBuiltinModule?.('node:v8') as
				| { getHeapStatistics?: () => { heap_size_limit?: number } }
				| undefined;
			const limit = valid(v8?.getHeapStatistics?.().heap_size_limit);
			if (limit) return limit;
		} catch {
			// A runtime without the module reports nothing.
		}
	}
	return valid(runtime.performance?.memory?.jsHeapSizeLimit);
}

/** The XML heap a read or edit may take on a host whose heap limit is `heapLimit` bytes. */
export function defaultXmlHeapBytes(heapLimit = runtimeHeapLimit() ?? ASSUMED_HEAP_LIMIT): number {
	return Math.max(
		MIN_BUDGET,
		Math.floor((heapLimit - HOST_HEAP_MARGIN) / PEAK_HEAP_PER_BUDGET_BYTE / MiB) * MiB,
	);
}
