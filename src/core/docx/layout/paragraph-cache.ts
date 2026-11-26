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

function copyLayout(result: ParagraphLayoutResult): ParagraphLayoutResult {
	// Inline objects, font fallbacks and frames can carry nested mutable data.
	if (
		result.frame ||
		result.lines.some((line) => line.fragments.some((frag) => frag.fontFallback || frag.object))
	)
		return structuredClone(result);
	return {
		...result,
		lines: result.lines.map((line) => ({
			...line,
			fragments: line.fragments.map((fragment) => ({ ...fragment })),
		})),
		pageBreakAfterLine: new Set(result.pageBreakAfterLine),
		columnBreakAfterLine: new Set(result.columnBreakAfterLine),
	};
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
		return copyLayout(cached.result);
	}
	const warnings: string[] = [];
	const result = compute((warning) => {
		warnings.push(warning);
		note(warning);
	});
	if (!cache.has(index) && cache.size >= 4000) cache.delete(cache.keys().next().value!);
	cache.set(index, { signature, result: copyLayout(result), warnings });
	return result;
}
