import type { VisioDocument, VisioShape } from 'ooxml-core/visio';

/** UTF-16 character bounds, applied before copying or case conversion. */
export const TEXT_SEARCH_LIMITS = Object.freeze({
	queryCharacters: 256,
	results: 500,
	indexedCharacters: 2_000_000,
	shapeCharacters: 32_768,
	previewCharacters: 160,
	shapes: 25_000,
	depth: 64,
});
export interface TextSearchResult {
	readonly pageIndex: number;
	readonly pageId: string;
	readonly pageName: string;
	readonly shapeId: string;
	readonly shapeName: string;
	/** Bounded start-of-text preview, not a match offset or highlighted excerpt. */
	readonly preview: string;
}
export interface TextSearchState {
	readonly query: string;
	readonly results: readonly TextSearchResult[];
	/** -1 until the user explicitly navigates, or after manual page/shape selection. */
	readonly activeIndex: number;
	readonly indexTruncated: boolean;
	readonly resultsTruncated: boolean;
}
interface SearchEntry {
	readonly result: TextSearchResult;
	readonly foldedText: string;
}
export interface DocumentTextIndex {
	readonly entries: readonly SearchEntry[];
	readonly truncated: boolean;
}
export const EMPTY_TEXT_SEARCH: TextSearchState = Object.freeze({
	query: '',
	results: Object.freeze([]),
	activeIndex: -1,
	indexTruncated: false,
	resultsTruncated: false,
});

/** Index canonical text for the current display state. No XML, DOM or layout parsing. */
export function indexDocumentText(
	model: VisioDocument | null,
	isVisible: (shape: VisioShape) => boolean = (shape) => !shape.hidden,
): DocumentTextIndex {
	const entries: SearchEntry[] = [];
	const seen = new Set<VisioShape>();
	let characters = 0,
		truncated = false;
	for (const [pageIndex, page] of (model?.pages ?? []).entries()) {
		const stack = page.shapes.map((shape) => ({ shape, depth: 0 })).reverse();
		while (stack.length) {
			const { shape, depth } = stack.pop()!;
			if (seen.has(shape) || depth > TEXT_SEARCH_LIMITS.depth) {
				truncated = true;
				continue;
			}
			if (seen.size >= TEXT_SEARCH_LIMITS.shapes)
				return Object.freeze({ entries: Object.freeze(entries), truncated: true });
			seen.add(shape);
			if (!isVisible(shape)) continue;
			if (!(shape.kind === 'group' && shape.groupDisplayMode === 0)) {
				const source = shape.text.plainText;
				const length = Math.min(
					source.length,
					TEXT_SEARCH_LIMITS.shapeCharacters,
					TEXT_SEARCH_LIMITS.indexedCharacters - characters,
				);
				if (length < source.length) truncated = true;
				if (length) {
					characters += length;
					const text = source.slice(0, length);
					const preview =
						text.slice(0, TEXT_SEARCH_LIMITS.previewCharacters) +
						(source.length > TEXT_SEARCH_LIMITS.previewCharacters ? '…' : '');
					entries.push(
						Object.freeze({
							foldedText: text.toLowerCase(),
							result: Object.freeze({
								pageIndex,
								pageId: page.id,
								pageName: page.name,
								shapeId: shape.id,
								shapeName: shape.name,
								preview,
							}),
						}),
					);
				}
			}
			for (let i = shape.children.length - 1; i >= 0; i--)
				stack.push({ shape: shape.children[i]!, depth: depth + 1 });
		}
	}
	return Object.freeze({ entries: Object.freeze(entries), truncated });
}

export function validateSearchQuery(query: string): void {
	if (typeof query !== 'string' || query.length > TEXT_SEARCH_LIMITS.queryCharacters)
		throw new Error(
			`Search queries must be strings of at most ${TEXT_SEARCH_LIMITS.queryCharacters} characters.`,
		);
}
/** One match per shape, literal ECMAScript lowercase comparison; no regex or language collation. */
export function searchDocumentText(index: DocumentTextIndex, query: string): TextSearchState {
	validateSearchQuery(query);
	const results: TextSearchResult[] = [];
	let resultsTruncated = false;
	if (query.trim()) {
		const folded = query.toLowerCase();
		for (const entry of index.entries) {
			if (!entry.foldedText.includes(folded)) continue;
			if (results.length === TEXT_SEARCH_LIMITS.results) {
				resultsTruncated = true;
				break;
			}
			results.push(entry.result);
		}
	}
	return Object.freeze({
		query,
		results: Object.freeze(results),
		activeIndex: -1,
		indexTruncated: index.truncated,
		resultsTruncated,
	});
}

export function textSearchStatus(search: TextSearchState): string {
	if (!search.query.trim()) return 'Find text across pages';
	const count = search.results.length;
	const result =
		search.activeIndex < 0
			? `${count}${search.resultsTruncated ? '+' : ''} matching shapes`
			: `${search.activeIndex + 1} of ${count}${search.resultsTruncated ? '+' : ''} matching shapes`;
	return result + (search.indexTruncated ? ' · Partial text index' : '');
}
