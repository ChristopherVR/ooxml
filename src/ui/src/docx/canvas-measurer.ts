import {
	createFakeMeasurer,
	cssFontStack,
	ligatureCss,
	type LayoutFontSpec,
	type TextMeasurer,
} from 'ooxml-core/docx/layout';

/**
 * Browser measurer backed by an offscreen `<canvas>` 2D context. Falls back to
 * the fake measurer's formula if no canvas is available (SSR, jsdom without
 * canvas support), so callers never need to branch on environment.
 */
export function createCanvasMeasurer(): TextMeasurer {
	let context: CanvasRenderingContext2D | null | undefined;
	const widthCache = new Map<string, number>();
	const metricCache = new Map<string, { ascent: number; height: number }>();
	const fallback = createFakeMeasurer();
	const ctx = (): CanvasRenderingContext2D | null => {
		if (context !== undefined) return context;
		try {
			context =
				typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
		} catch {
			context = null;
		}
		return context;
	};
	const fontString = (font: LayoutFontSpec) =>
		`${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${font.sizePx}px ${cssFontStack(font.family)}`;
	const metricsOf = (c: CanvasRenderingContext2D, font: LayoutFontSpec) => {
		const key = fontString(font);
		const cached = metricCache.get(key);
		if (cached) return cached;
		c.font = key;
		const measured = c.measureText('Mg');
		const metrics = {
			ascent: measured.fontBoundingBoxAscent ?? font.sizePx * 0.95,
			height:
				measured.fontBoundingBoxAscent !== undefined &&
				measured.fontBoundingBoxDescent !== undefined
					? measured.fontBoundingBoxAscent + measured.fontBoundingBoxDescent
					: font.sizePx * 1.22,
		};
		if (metricCache.size >= 256) metricCache.delete(metricCache.keys().next().value!);
		metricCache.set(key, metrics);
		return metrics;
	};
	return {
		widthOf(text, font) {
			const c = ctx();
			if (!c) return fallback.widthOf(text, font);
			const key = `${fontString(font)}\u0000${font.kerning ?? 'auto'}\u0000${font.ligatures ?? 'auto'}\u0000${font.smallCaps ?? false}\u0000${text}`;
			const cached = widthCache.get(key);
			if (cached !== undefined) return cached;
			c.font = fontString(font);
			c.fontKerning = font.kerning ?? 'auto';
			let width = c.measureText(text).width;
			// Canvas has no ligature feature selector. Measure explicit selections with the same
			// browser shaping used to draw Print Layout, before applying glyph scale and spacing.
			if ((font.ligatures !== undefined || font.smallCaps) && document.body) {
				const span = document.createElement('span');
				span.style.cssText =
					'position:fixed;left:-100000px;top:0;visibility:hidden;display:inline-block;white-space:pre;padding:0;border:0;margin:0;letter-spacing:0;text-transform:none';
				span.style.font = fontString(font);
				span.style.fontKerning = font.kerning ?? 'auto';
				if (font.ligatures !== undefined)
					span.style.fontVariantLigatures = ligatureCss(font.ligatures);
				if (font.smallCaps) span.style.fontVariantCaps = 'small-caps';
				span.textContent = text;
				document.body.append(span);
				try {
					width = span.getBoundingClientRect().width;
				} finally {
					span.remove();
				}
			}
			if (widthCache.size >= 20000) widthCache.delete(widthCache.keys().next().value!);
			widthCache.set(key, width);
			return width;
		},
		ascentOf(font) {
			const c = ctx();
			if (!c) return fallback.lineHeightOf(font) * 0.8;
			return metricsOf(c, font).ascent;
		},
		lineHeightOf(font) {
			const c = ctx();
			if (!c) return fallback.lineHeightOf(font);
			return metricsOf(c, font).height;
		},
	};
}
