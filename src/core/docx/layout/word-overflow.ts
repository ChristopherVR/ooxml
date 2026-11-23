const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

/** Take the longest fitting prefix, allowing one indivisible grapheme to overflow if necessary. */
export function fittingWordPrefix(
	word: string,
	width: number,
	measure: (text: string) => number,
): string {
	const segments = segmenter.segment(word);
	const endOf = (index: number) => {
		const segment = segments.containing(index)!;
		return segment.index + segment.segment.length;
	};
	let best = endOf(0);
	// Find an upper bound near the line width before searching. Measuring half of
	// a very long remaining URL on every line would repeatedly shape the same text.
	let high = best;
	while (high < word.length) {
		const end = endOf(high);
		if (measure(word.slice(0, end)) > width) break;
		best = end;
		high = Math.min(word.length, end * 2);
	}
	high = Math.min(high, word.length - 1);
	while (best <= high) {
		const probe = endOf(Math.floor((best + high) / 2));
		if (measure(word.slice(0, probe)) <= width) {
			best = probe;
		} else {
			high = segments.containing(probe - 1)!.index - 1;
		}
	}
	return word.slice(0, best);
}
