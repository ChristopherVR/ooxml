/** What a status bar shows at a given width; the product decides what each part is. */
export interface StatusLayout {
	/** The word count (lowest priority text after the language). */
	words: boolean;
	/** The proofing language. */
	language: boolean;
	/** The view switcher. */
	views: boolean;
	/** Show the page counter as "1/3" (the full text moves to its tooltip). */
	compactPage: boolean;
	/** How many selection aggregates (sum, average, count) fit. */
	stats: number;
}

/**
 * Office drops status items from least to most important as the bar narrows: extra aggregates, the
 * language, the view switcher, the word count, then the last aggregate. The page or cell mode, the
 * notes button and the zoom control always stay. A width of 0 (not measured yet) shows everything.
 */
export function statusLayout(width: number): StatusLayout {
	if (!(width > 0))
		return { words: true, language: true, views: true, compactPage: false, stats: Infinity };
	return {
		words: width >= 440,
		language: width >= 640,
		views: width >= 480,
		compactPage: width < 520,
		stats: width >= 760 ? Infinity : width >= 560 ? 2 : width >= 440 ? 1 : 0,
	};
}

/** Calls `onChange` with the bar's width whenever it changes; returns a stop function. */
export function watchStatusWidth(element: HTMLElement, onChange: (width: number) => void) {
	if (typeof ResizeObserver === 'undefined') return () => {};
	let last = -1;
	const observer = new ResizeObserver(() => {
		const width = element.clientWidth;
		if (width === last) return;
		last = width;
		onChange(width);
	});
	observer.observe(element);
	return () => observer.disconnect();
}

/** "Page 2 of 5" as "2/5" in any locale: its numbers only; text without two numbers is kept. */
export function compactPageText(text: string): string {
	const numbers = text.match(/\d+/g);
	return numbers && numbers.length === 2 ? numbers.join('/') : text;
}
