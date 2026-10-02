// The browser text measurer (the one piece of layout that needs a DOM): an offscreen canvas 2D
// context with a width cache, injected into the core's overflow and auto-fit functions. Without a
// canvas (SSR, jsdom) it falls back to an average-character estimate so callers never branch.
import type { FontView } from '@christophervr/xlsx-core';
import { cssFont } from './cell-paint.js';

export interface TextMeasurer {
	/** Width in CSS pixels of `text` in a CSS font shorthand. */
	measure(text: string, font: string): number;
	/** Width at 100% zoom of `text` in a core font view (for `autoFitColumnWidth`). */
	measureFont(text: string, font: FontView): number;
}

const sizeOf = (font: string): number => Number(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? 14.67);

export function estimateWidth(text: string, font: string): number {
	const size = sizeOf(font);
	const bold = /\bbold\b/.test(font) ? 1.07 : 1;
	let units = 0;
	for (const ch of text)
		units += /[ilI.,:;'|!]/.test(ch)
			? 0.28
			: /[mwMW@]/.test(ch)
				? 0.85
				: /[　-鿿]/.test(ch)
					? 1
					: 0.52;
	return units * size * bold;
}

export function createTextMeasurer(doc: Document | undefined): TextMeasurer {
	let context: CanvasRenderingContext2D | null | undefined;
	const cache = new Map<string, number>();
	const ctx = (): CanvasRenderingContext2D | null => {
		if (context !== undefined) return context;
		try {
			const isJsdom = typeof navigator !== 'undefined' && /jsdom/i.test(navigator.userAgent);
			context = !doc || isJsdom ? null : doc.createElement('canvas').getContext('2d');
		} catch {
			context = null;
		}
		return context;
	};
	const measure = (text: string, font: string): number => {
		if (!text) return 0;
		const key = `${font}\u0000${text}`;
		const hit = cache.get(key);
		if (hit !== undefined) return hit;
		const c = ctx();
		let width: number;
		if (c) {
			if (c.font !== font) c.font = font;
			width = c.measureText(text).width;
		} else width = estimateWidth(text, font);
		if (cache.size > 50_000) cache.clear();
		cache.set(key, width);
		return width;
	};
	return {
		measure,
		measureFont: (text, font) => measure(text, cssFont(font, 100)),
	};
}
