import {
	createFakeMeasurer,
	cssFontStack,
	ligatureCss,
	type LayoutFontSpec,
	type TextMeasurer,
} from 'ooxml-core/docx/layout';
import { typoLineGapEm } from 'ooxml-core/text';

/** The canvas measurer, plus where Print Layout must move each font's text to match its ascent. */
export interface CanvasMeasurer extends TextMeasurer {
	/**
	 * How far (CSS px) below the baseline the browser draws in a box of the font's line height
	 * (half the leading above) the measured ascent puts it: half the leading for a font without
	 * `USE_TYPO_METRICS`, whose leading Word places wholly above the text, else 0.
	 */
	baselineShiftOf(font: LayoutFontSpec): number;
}

/**
 * Browser measurer backed by an offscreen `<canvas>` 2D context. Falls back to
 * the fake measurer's formula if no canvas is available (SSR, jsdom without
 * canvas support), so callers never need to branch on environment.
 */
export function createCanvasMeasurer(): CanvasMeasurer {
	let context: CanvasRenderingContext2D | null | undefined;
	const widthCache = new Map<string, number>();
	const metricCache = new Map<string, { ascent: number; height: number; shift: number }>();
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
	// The browser rounds a font's ascender, descender and line box to whole pixels. At 100 px that
	// cost Century Gothic 0.6% of its line height (1.22 em for Word's 1.2261), enough to fit one
	// more line on a page than Word (#383). 2048 px is exact for 2048-unit TrueType fonts and
	// within 1/4096 em for the rest.
	const referencePx = 2048;
	const fontString = (font: LayoutFontSpec, sizePx = font.sizePx) =>
		`${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${sizePx}px ${cssFontStack(font.family, font.fallback)}`;
	const metricsOf = (c: CanvasRenderingContext2D, font: LayoutFontSpec) => {
		// Measure a reference size once per face. Canvas rounds the ascender and descender
		// separately and doesn't include the font's line gap; the span's line box does. That gap is
		// the typographic one for a font with USE_TYPO_METRICS, split above and below the text, and
		// otherwise the hhea leading the Windows metrics leave out, which Word puts all above the
		// ascent (#382: Times New Roman and Arial first lines sat up to 1.6 pt high with it split).
		const key = fontString(font, referencePx);
		const cached = metricCache.get(key);
		if (cached) return cached;
		c.font = key;
		const measured = c.measureText('Mg');
		const ascent = measured.fontBoundingBoxAscent ?? referencePx * 0.95;
		const fontBox = ascent + (measured.fontBoundingBoxDescent ?? referencePx * 0.27);
		const span = document.createElement('span');
		span.style.cssText =
			'position:fixed;visibility:hidden;display:inline-block;padding:0;border:0;margin:0';
		span.style.font = key;
		span.style.lineHeight = 'normal';
		span.textContent = 'Mg';
		document.body?.append(span);
		let height: number;
		try {
			height = span.getBoundingClientRect().height || fontBox;
		} finally {
			span.remove();
		}
		const gap = height - fontBox;
		const above = typoLineGapEm(font.family) > 0 ? gap / 2 : gap;
		const metrics = {
			ascent: (ascent + above) / referencePx,
			height: height / referencePx,
			shift: (above - gap / 2) / referencePx,
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
			return metricsOf(c, font).ascent * font.sizePx;
		},
		lineHeightOf(font) {
			const c = ctx();
			if (!c) return fallback.lineHeightOf(font);
			return metricsOf(c, font).height * font.sizePx;
		},
		baselineShiftOf(font) {
			const c = ctx();
			return c ? metricsOf(c, font).shift * font.sizePx : 0;
		},
	};
}
