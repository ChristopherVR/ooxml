import type { Node } from 'prosemirror-model';

const segmenters = new Map<string, Intl.Segmenter>();
const documentCounts = new WeakMap<Node, Map<string, number>>();

/** Unicode-aware word counting; segmentation is provided by the host's ICU data. */
export function countWords(text: string, locale?: string): number {
	if (typeof Intl.Segmenter === 'function') {
		const key = locale || '';
		let segmenter = segmenters.get(key);
		if (!segmenter) {
			try {
				segmenter = new Intl.Segmenter(locale || undefined, { granularity: 'word' });
			} catch {
				segmenter = new Intl.Segmenter(undefined, { granularity: 'word' });
			}
			if (segmenters.size >= 16) segmenters.clear();
			segmenters.set(key, segmenter);
		}
		let count = 0;
		for (const part of segmenter.segment(text)) if (part.isWordLike) count++;
		return count;
	}
	// Older engines cannot provide dictionary segmentation for unspaced scripts.
	return text.match(/[\p{L}\p{N}][\p{L}\p{M}\p{N}'’]*/gu)?.length ?? 0;
}

/** Immutable editor subtrees retain their counts across typing, selection changes and undo. */
export function documentWordCount(node: Node, locale?: string): number {
	const key = locale || '';
	const cached = documentCounts.get(node);
	const value = cached?.get(key);
	if (value !== undefined) return value;
	let count = 0;
	if (node.isTextblock) count = countWords(node.textBetween(0, node.content.size, ' '), locale);
	else node.forEach((child) => (count += documentWordCount(child, locale)));
	const counts = cached ?? new Map<string, number>();
	counts.set(key, count);
	documentCounts.set(node, counts);
	return count;
}
