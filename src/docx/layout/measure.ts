/** A font as the measurer needs it; sizes are already in CSS pixels. */
export interface LayoutFontSpec {
	family: string;
	sizePx: number;
	bold?: boolean;
	italic?: boolean;
	kerning?: 'normal' | 'none';
	ligatures?: import('../ligatures.js').Ligatures;
	smallCaps?: boolean;
}

/** Injectable text measurer: a host supplies canvas or font tables, tests use the fake below. The engine never touches the DOM. */
export interface TextMeasurer {
	/** Glyph width in CSS pixels, including the requested kerning mode, before scale/spacing. */
	widthOf(text: string, font: LayoutFontSpec): number;
	/** Natural single-spaced line height for `font`, in CSS pixels (ascent + descent + leading). */
	lineHeightOf(font: LayoutFontSpec): number;
	/** Height above the baseline within that line height; about 80% when not provided. */
	ascentOf?(font: LayoutFontSpec): number;
}

/** A font's ascent and descent (together its single line height), for baseline alignment. */
export function fontMetrics(
	measurer: TextMeasurer,
	font: LayoutFontSpec,
): { ascent: number; descent: number } {
	const height = measurer.lineHeightOf(font);
	const ascent = Math.min(height, measurer.ascentOf?.(font) ?? height * 0.8);
	return { ascent, descent: height - ascent };
}

/**
 * Deterministic measurer for unit tests: fixed per-character advance and line
 * height derived only from font size, so line-breaking and pagination tests
 * never depend on host font metrics or canvas availability.
 */
export function createFakeMeasurer(
	options: { charWidthFactor?: number; lineHeightFactor?: number } = {},
): TextMeasurer {
	const charWidthFactor = options.charWidthFactor ?? 0.6;
	const lineHeightFactor = options.lineHeightFactor ?? 1.15;
	return {
		widthOf(text, font) {
			let width = 0;
			for (const ch of text) {
				if (ch === '\t' || ch === '\n') continue;
				width += font.sizePx * charWidthFactor * (font.bold ? 1.08 : 1) * (ch === ' ' ? 0.6 : 1);
			}
			return width;
		},
		lineHeightOf(font) {
			return font.sizePx * lineHeightFactor;
		},
	};
}
