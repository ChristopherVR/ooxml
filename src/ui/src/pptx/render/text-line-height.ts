/**
 * `line-height` resolution shared by every binding's text renderer: PowerPoint's
 * single-spacing pitch, the proportional (`a:spcPct`) / exact (`a:spcPts`)
 * spacing modes, `compatLnSpc`'s legacy model, and converting a resolved
 * `line-height` CSS value back to a px measurement.
 *
 * Split out of `text-style-helpers.ts` to keep that module focused and under
 * the repo's ~300 LOC guideline.
 */

/** Minimal line-spacing fields {@link resolveLineHeight} needs. */
export interface LineHeightSource {
	lineSpacing?: number;
	lineSpacingExactPt?: number;
	/**
	 * `a:bodyPr/@compatLnSpc`: use PowerPoint's legacy (97-2003) line-spacing
	 * model instead of the newer "exact" one. See {@link proportionalLineHeight}
	 * for what this changes.
	 */
	compatibleLineSpacing?: boolean;
}

/**
 * PowerPoint's single-spacing line height as a unitless multiplier of the font
 * size. Measured against PowerPoint itself (COM `TextRange2.BoundHeight` on
 * the issue #131 deck): every single-spaced Arial paragraph lays out at
 * exactly 1.2x its point size (10.5pt -> 12.6pt, 10pt -> 12pt, 8pt -> 9.6pt),
 * and the browser default `normal` (~1.25 for Arial) accumulated a visible
 * vertical drift over a full text panel.
 */
export const DEFAULT_LINE_HEIGHT = 1.2;

/**
 * Resolve a CSS `line-height` value from a TextStyle's spacing fields.
 *
 * - If `lineSpacingExactPt` is set (exact point mode from `a:lnSpc > a:spcPts`),
 *   returns a fixed `"<n>pt"` string.
 * - Otherwise the proportional multiplier from `a:spcPct` (`lineSpacing`) is
 *   used, defaulting to PowerPoint's single-spacing {@link DEFAULT_LINE_HEIGHT}.
 *
 * Returning a unitless multiplier (rather than relying on the browser's
 * font-dependent `normal`, which is ~1.2-1.5) lets the value scale with the
 * resolved font size, keeping multi-line text inside its box.
 *
 * @param textStyle     The text style carrying the spacing fields (may be
 *                      `undefined`).
 * @param hasItalicRuns Whether the block contains italic runs. No longer
 *                      changes the default: PowerPoint does not add leading
 *                      for italics, and the old 1.35 bump pushed every italic
 *                      block visibly below its authored position.
 */
export function resolveLineHeight(
	textStyle: LineHeightSource | undefined,
	hasItalicRuns: boolean,
): string | number {
	if (typeof textStyle?.lineSpacingExactPt === 'number' && textStyle.lineSpacingExactPt > 0) {
		return `${textStyle.lineSpacingExactPt}pt`;
	}
	void hasItalicRuns;
	// `a:spcPct` multiplies PowerPoint's single-spacing pitch, it does not
	// replace it: 200% at 20pt lays out at a 48pt pitch (2.4x, measured via COM
	// export on the issue #132 deck), so the multiplier stacks on the 1.2 base.
	// `compatLnSpc` opts a body out of that stacking (see
	// `proportionalLineHeight`'s doc comment).
	return proportionalLineHeight(textStyle?.lineSpacing, textStyle?.compatibleLineSpacing);
}

/**
 * CSS line-height for a proportional `a:spcPct` multiplier: the multiplier
 * stacked on PowerPoint's {@link DEFAULT_LINE_HEIGHT} single-spacing pitch,
 * or that base itself when no multiplier is set.
 *
 * `compatLineSpacing` (`a:bodyPr/@compatLnSpc`) changes nothing here. It was
 * once modelled as "the multiplier alone, no 1.2 pitch", which was never
 * measured and painted every body of the older Office themes (whose master
 * sets `compatLnSpc="1"`) about 17% too tight. PowerPoint 16 static exports
 * of Calibri, Arial and Verdana at 28pt with `compatLnSpc` 0 and 1 and
 * `lnSpc` unset, 90% and 150% all lay out at 1.2x the font size times the
 * multiplier with the flag on (44.8, 40.3 and 67.2px at 37.3px). The
 * parameter stays so callers keep their signature.
 */
export function proportionalLineHeight(
	lineSpacing: number | undefined,
	compatLineSpacing = false,
): number {
	void compatLineSpacing;
	return typeof lineSpacing === 'number' && lineSpacing > 0
		? lineSpacing * DEFAULT_LINE_HEIGHT
		: DEFAULT_LINE_HEIGHT;
}

/**
 * Resolve a single line's height in px from a `line-height` CSS value (as
 * {@link resolveLineHeight} produces): a bare unitless multiplier (`×
 * fontSizePx`), a `"<n>pt"` / `"<n>px"` string, or a fallback to PowerPoint's
 * single-spacing pitch when the value is neither.
 */
export function lineHeightToPx(
	fontSizePx: number,
	lineHeight: string | number | undefined,
): number {
	if (typeof lineHeight === 'number') {
		return fontSizePx * lineHeight;
	}
	if (typeof lineHeight === 'string') {
		const match = /^([\d.]+)(pt|px)$/u.exec(lineHeight);
		if (match) {
			const value = Number.parseFloat(match[1]);
			return match[2] === 'pt' ? value * (96 / 72) : value;
		}
	}
	return fontSizePx * DEFAULT_LINE_HEIGHT;
}
