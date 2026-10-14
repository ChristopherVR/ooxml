import type { TextMeasurer } from './measure';
import type { ParagraphLayoutResult } from './paragraph-layout';

interface Entry {
	signature: string;
	result: ParagraphLayoutResult;
	warnings: string[];
}

// Only explicit layout sessions opt in. Arbitrary public measurers may change over time.
const sessions = new WeakMap<TextMeasurer, Map<string, Entry>>();

export function registerParagraphCache(measurer: TextMeasurer): void {
	sessions.set(measurer, new Map());
}

export function cachedParagraphLayout(
	measurer: TextMeasurer,
	key: () => { index: string; signature: string },
	note: (message: string) => void,
	compute: (note: (message: string) => void) => ParagraphLayoutResult,
): ParagraphLayoutResult {
	const cache = sessions.get(measurer);
	if (!cache) return compute(note);
	const { index, signature } = key();
	const cached = cache.get(index);
	if (cached?.signature === signature) {
		for (const warning of cached.warnings) note(warning);
		// Placement and callers own mutable results, never the cache's copy.
		return structuredClone(cached.result);
	}
	const warnings: string[] = [];
	const result = compute((warning) => {
		warnings.push(warning);
		note(warning);
	});
	if (!cache.has(index) && cache.size >= 4000) cache.delete(cache.keys().next().value!);
	cache.set(index, { signature, result: structuredClone(result), warnings });
	return result;
}
