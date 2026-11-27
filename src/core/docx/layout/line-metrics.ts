import { fontMetrics, type LayoutFontSpec, type TextMeasurer } from './measure';
import type { BreakToken } from './text-breaks';
import type { LayoutParagraph } from './input';
import { DEFAULT_FONT_SIZE_PT, ptToPx, twipsToPx } from './units';
import { fontOf, type PlacedToken } from './paragraph-tokens';
import { contentControlTagMetrics } from './content-control-tag-metrics';

/** How far superscript text is raised and subscript lowered, as a share of the run's size. */
const SUPER_RAISE = 0.33;
const SUB_DROP = 0.14;

const tagExtent = {
	above: ptToPx(contentControlTagMetrics.lineAscentPt),
	below: ptToPx(contentControlTagMetrics.lineDescentPt),
	ascent: ptToPx(contentControlTagMetrics.boxAscentPt),
	descent: ptToPx(contentControlTagMetrics.boxHeightPt - contentControlTagMetrics.boxAscentPt),
};

/** Where a token's box sits relative to the baseline: its extent above and below it. */
export function tokenExtent(
	token: BreakToken,
	paragraph: LayoutParagraph,
	fonts: LayoutFontSpec[],
	measurer: TextMeasurer,
): { above: number; below: number; ascent: number; descent: number } {
	const run = paragraph.runs[token.runIndex];
	if (token.kind === 'object' && run?.object) {
		const below = Math.min(run.object.descentPx ?? 0, run.object.heightPx);
		const above = run.object.heightPx - below;
		return { above, below, ascent: above, descent: below };
	}
	if (run?.contentControlTag) return tagExtent;
	const { ascent, descent } = fontMetrics(measurer, fonts[token.runIndex] ?? fontOf({ text: '' }));
	const size = ptToPx(run?.fontSizePt ?? DEFAULT_FONT_SIZE_PT);
	const scriptShift =
		run?.script === 'super' ? SUPER_RAISE * size : run?.script === 'sub' ? -SUB_DROP * size : 0;
	const shift = scriptShift + (run?.positionPx ?? 0);
	return { above: ascent + shift, below: descent - shift, ascent, descent };
}

/**
 * A line's height and baseline: the tallest extent above and below the baseline over its tokens
 * (inline pictures sit on the baseline), scaled by the paragraph's line spacing. Exact and
 * at-least spacing add their extra space above the text, so the baseline sits one descent above
 * the line's bottom. Multiple spacing above single (`w:lineRule="auto"`, `w:line` > 240) adds it
 * below the text instead: Word keeps the baseline one ascent below the line's top and lets that
 * extra space (`extraBelowPx`) of a page's last line run past the bottom margin (#360).
 */
export function lineMetrics(
	placed: PlacedToken[],
	fonts: LayoutFontSpec[],
	paragraph: LayoutParagraph,
	measurer: TextMeasurer,
): { heightPx: number; baselinePx: number; extraBelowPx: number } {
	const empty = placed.length
		? undefined
		: fontMetrics(measurer, paragraph.emptyLineFont ?? fonts[0] ?? fontOf({ text: '' }));
	let above = empty?.ascent ?? 0;
	let below = empty?.descent ?? 0;
	let objectAscent = 0;
	let objectDescent = 0;
	let textSizePx = 0;
	let marker: { above: number; below: number; sizePx: number } | undefined;
	for (const { token } of placed) {
		const run = paragraph.runs[token.runIndex];
		const metrics = run?.object?.noteSeparator
			? fontMetrics(measurer, fonts[token.runIndex] ?? fontOf(run))
			: undefined;
		const extent = metrics
			? { above: metrics.ascent, below: metrics.descent }
			: tokenExtent(token, paragraph, fonts, measurer);
		const sizePx = fonts[token.runIndex]?.sizePx ?? 0;
		if (run?.marker) {
			marker = { ...extent, sizePx };
			continue;
		}
		if (token.kind === 'object' && !metrics) {
			objectAscent = Math.max(objectAscent, extent.above);
			objectDescent = Math.max(objectDescent, extent.below);
		} else {
			textSizePx = Math.max(textSizePx, sizePx);
		}
		above = Math.max(above, extent.above);
		below = Math.max(below, extent.below);
	}
	// A list label that raises the baseline uses its own descent at the text size.
	// An oversized label adds room above the text without increasing its descent.
	if (marker && marker.above > above) {
		const markerDescent = textSizePx && marker.sizePx > textSizePx ? below : marker.below;
		below = Math.max(markerDescent, objectDescent, above + below - marker.above);
		above = marker.above;
	}
	const natural = above + below;
	const rule = paragraph.lineSpacingRule ?? 'auto';
	const spacingTwips = paragraph.lineSpacingTwips;
	if (rule !== 'exact' && rule !== 'atLeast') {
		const heightPx = Math.max(objectAscent + below, natural * ((spacingTwips ?? 240) / 240));
		const extraBelowPx = Math.max(0, heightPx - natural);
		return { heightPx, baselinePx: heightPx - below - extraBelowPx, extraBelowPx };
	}
	const heightPx =
		rule === 'exact'
			? spacingTwips != null
				? twipsToPx(spacingTwips)
				: natural
			: Math.max(natural, spacingTwips != null ? twipsToPx(spacingTwips) : 0);
	return { heightPx, baselinePx: heightPx - below, extraBelowPx: 0 };
}
