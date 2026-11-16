/**
 * Standalone parser utilities for `a:pPr` (paragraph properties) XML nodes.
 *
 * The attributes, spacing, bullet group and tab stops are read by the shared
 * DrawingML codec (`drawingml/text-paragraph-codec.ts`, through
 * `paragraph-properties-neutral`), the same rules the DOM reader runs; these
 * functions convert the neutral values into the pptx {@link TextStyle} and
 * {@link BulletInfo} fields (EMU and points to px, alignment tokens, the
 * PowerPoint line-spacing clamp and the spacing basis).
 *
 * @module paragraph-properties-parser
 */

import type { DrawingParagraphProperties } from '../../../drawingml/text-model';
import type { DrawingTextSpacing } from '../../../drawingml/types';
import {
	readDrawingBulletChild,
	readDrawingTextSpacing,
} from '../../../drawingml/text-paragraph-codec';
import type { BulletInfo, TextStyle, XmlObject } from '../types';
import {
	objectParagraphAttributes,
	objectParagraphProperties,
	objectTextElementView,
} from './paragraph-properties-neutral';

const EMU_PER_PX = 9525;

function pointsToPixels(points: number): number {
	return points * (96 / 72);
}

// ---------------------------------------------------------------------------
// Alignment
// ---------------------------------------------------------------------------

const ALIGN_MAP: Record<string, TextStyle['align']> = {
	l: 'left',
	ctr: 'center',
	r: 'right',
	just: 'justify',
	justify: 'justify',
	justLow: 'justLow',
	dist: 'dist',
	thaiDist: 'thaiDist',
};

/**
 * Map an OOXML `a:pPr/@algn` token to its {@link TextStyle.align} value.
 */
export function parseAlignmentAttr(algn: string | undefined): TextStyle['align'] | undefined {
	if (!algn) {
		return undefined;
	}
	return ALIGN_MAP[algn] || undefined;
}

/**
 * Resolve a single paragraph's effective alignment: its own `a:pPr/@algn`
 * wins outright; otherwise the alignment inherited from its OWN placeholder
 * level (not a sibling paragraph's); otherwise the rtl-based default.
 *
 * Callers must look up `placeholderAlignment` per paragraph (the level the
 * CURRENT paragraph's own `lvl` resolves to) rather than reusing a
 * shape-level snapshot: a snapshot taken from an earlier paragraph in the
 * same shape masks this paragraph's own placeholder-level alignment once
 * that earlier paragraph declared an explicit `algn` (for example a centered
 * heading followed by left-aligned body bullets).
 */
export function resolveParagraphAlignment(
	authoredAlignment: unknown,
	placeholderAlignment: TextStyle['align'] | undefined,
	paragraphRtl: boolean | undefined,
): TextStyle['align'] {
	if (authoredAlignment !== undefined && authoredAlignment !== null) {
		return parseAlignmentAttr(String(authoredAlignment)) ?? 'left';
	}
	return placeholderAlignment ?? (paragraphRtl ? 'right' : 'left');
}

// ---------------------------------------------------------------------------
// Spacing helpers
// ---------------------------------------------------------------------------

/**
 * A single line's height, as a multiple of its font's point size: the basis
 * `a:spcPct` paragraph spacing (`a:spcBef`/`a:spcAft`) resolves against.
 */
const SINGLE_LINE_HEIGHT_FACTOR = 1.2;

/**
 * Parse paragraph spacing from an `a:spcBef` or `a:spcAft` node.
 *
 * `a:spcPts/@_val` (hundredths of a point) converts to pixels at 96 dpi.
 * `a:spcPct` is a fraction of one single-spaced line, so it resolves only when
 * the paragraph's font size in px is given. Returns `undefined` when
 * absent/invalid.
 *
 * A single line is 1.2x the font size, not 1.0x: measured through COM on a
 * slide with 18pt runs and `spcBef` 100% / `spcAft` 50% (audit-text/gen.py
 * slide 16), the pitch between consecutive paragraph baselines was exactly
 * `(1.0 + 1.2 + 0.5 * 1.2) * 18pt`, i.e. the paragraph's own single-line pitch
 * plus `spcAft` and `spcBef` each scaled by `1.2 * fontSize`.
 */
export function parseParagraphSpacingPx(
	spacingNode: XmlObject | undefined,
	basisFontSizePx?: number,
): number | undefined {
	return paragraphSpacingPxFromDrawing(
		readDrawingTextSpacing(objectTextElementView(spacingNode)),
		basisFontSizePx,
	);
}

/** A neutral `a:spcBef` / `a:spcAft` value in px (see {@link parseParagraphSpacingPx}). */
export function paragraphSpacingPxFromDrawing(
	spacing: DrawingTextSpacing | undefined,
	basisFontSizePx?: number,
): number | undefined {
	if (spacing?.unit === 'points') {
		return pointsToPixels(spacing.value);
	}
	if (spacing && basisFontSizePx !== undefined && basisFontSizePx > 0) {
		return spacing.value * basisFontSizePx * SINGLE_LINE_HEIGHT_FACTOR;
	}
	return undefined;
}

/**
 * Parse `a:buSzPct/@_val` into a bullet size expressed as a percentage of the
 * run text (so `25000` and `25%` both yield `25`).
 *
 * @param buSzPct - The `a:buSzPct` node, when the level declares one.
 * @returns Percentage of the run size, or `undefined` when absent/unparseable.
 */
export function parseBulletSizePercent(buSzPct: XmlObject | undefined): number | undefined {
	const view = objectTextElementView(buSzPct);
	const out: DrawingParagraphProperties = {};
	if (!view || !readDrawingBulletChild('buSzPct', view, out)) return undefined;
	const size = out.bulletSize;
	return typeof size === 'object' ? size.value * 100 : undefined;
}

/** The PowerPoint clamp of a proportional line spacing multiplier. */
export function clampLineSpacingMultiplier(multiplier: number): number {
	return Math.max(0.1, Math.min(5, multiplier));
}

/**
 * Parse proportional line spacing from `a:lnSpc > a:spcPct/@_val`.
 *
 * Accepts both lexical forms of the value (`150000` and `150%`). Returns a
 * multiplier clamped to [0.1, 5].
 */
export function parseLineSpacingMultiplier(
	lineSpacingNode: XmlObject | undefined,
): number | undefined {
	const spacing = readDrawingTextSpacing(objectTextElementView(lineSpacingNode));
	return spacing?.unit === 'percent' ? clampLineSpacingMultiplier(spacing.value) : undefined;
}

/**
 * Parse exact line spacing from `a:lnSpc > a:spcPts/@_val`.
 *
 * The OOXML value is in hundredths of a point.  Returns the value in points.
 */
export function parseLineSpacingExactPt(
	lineSpacingNode: XmlObject | undefined,
): number | undefined {
	const spacing = readDrawingTextSpacing(objectTextElementView(lineSpacingNode));
	return spacing?.unit === 'points' && spacing.value > 0 ? spacing.value : undefined;
}

// ---------------------------------------------------------------------------
// Margins and indentation
// ---------------------------------------------------------------------------

/**
 * Parse paragraph margins and indent from `a:pPr` attributes.
 *
 * `@_marL`, `@_marR` are in EMU; `@_indent` is in EMU (may be negative
 * for hanging indent).
 */
export function parseParagraphMargins(
	pPr: XmlObject | undefined,
): Pick<TextStyle, 'paragraphMarginLeft' | 'paragraphMarginRight' | 'paragraphIndent'> {
	return paragraphMarginsFromDrawing(objectParagraphAttributes(pPr));
}

/** The margin fields of neutral paragraph properties, in px. */
export function paragraphMarginsFromDrawing(
	properties: DrawingParagraphProperties,
): Pick<TextStyle, 'paragraphMarginLeft' | 'paragraphMarginRight' | 'paragraphIndent'> {
	const result: Pick<
		TextStyle,
		'paragraphMarginLeft' | 'paragraphMarginRight' | 'paragraphIndent'
	> = {};
	if (properties.marginLeftEmu !== undefined) {
		result.paragraphMarginLeft = properties.marginLeftEmu / EMU_PER_PX;
	}
	if (properties.marginRightEmu !== undefined) {
		result.paragraphMarginRight = properties.marginRightEmu / EMU_PER_PX;
	}
	if (properties.indentEmu !== undefined) {
		result.paragraphIndent = properties.indentEmu / EMU_PER_PX;
	}
	return result;
}

// ---------------------------------------------------------------------------
// RTL + level
// ---------------------------------------------------------------------------

/**
 * Parse RTL direction from `a:pPr/@_rtl`.
 */
export function parseParagraphRtl(pPr: XmlObject | undefined): boolean | undefined {
	return objectParagraphAttributes(pPr).rightToLeft;
}

/**
 * Parse paragraph level from `a:pPr/@_lvl`.
 * Returns 0-8 (clamped).
 */
export function parseParagraphLevel(pPr: XmlObject | undefined): number {
	return objectParagraphAttributes(pPr).level ?? 0;
}

// ---------------------------------------------------------------------------
// Tab stops (implemented in `tab-stops`; re-exported for existing callers)
// ---------------------------------------------------------------------------

export { parseTabStops, serializeTabStop } from './tab-stops';

// ---------------------------------------------------------------------------
// Additional paragraph attributes
// ---------------------------------------------------------------------------

type ExtraAttributes = Pick<
	TextStyle,
	'defaultTabSize' | 'eaLineBreak' | 'latinLineBreak' | 'fontAlignment' | 'hangingPunctuation'
>;

/**
 * Parse additional paragraph-level attributes from `a:pPr`.
 *
 * Handles: `@_defTabSz`, `@_eaLnBrk`, `@_latinLnBrk`, `@_fontAlgn`,
 * `@_hangingPunct`.
 */
export function parseParagraphExtraAttributes(pPr: XmlObject | undefined): ExtraAttributes {
	return paragraphExtraAttributesFromDrawing(objectParagraphAttributes(pPr));
}

/** The additional attribute fields of neutral paragraph properties. */
export function paragraphExtraAttributesFromDrawing(
	properties: DrawingParagraphProperties,
): ExtraAttributes {
	const result: ExtraAttributes = {};
	if (properties.defaultTabSizeEmu !== undefined) {
		result.defaultTabSize = properties.defaultTabSizeEmu / EMU_PER_PX;
	}
	if (properties.eastAsianLineBreak !== undefined) {
		result.eaLineBreak = properties.eastAsianLineBreak;
	}
	if (properties.latinLineBreak !== undefined) {
		result.latinLineBreak = properties.latinLineBreak;
	}
	if (properties.fontAlign !== undefined) {
		result.fontAlignment = properties.fontAlign;
	}
	if (properties.hangingPunctuation !== undefined) {
		result.hangingPunctuation = properties.hangingPunctuation;
	}
	return result;
}

// ---------------------------------------------------------------------------
// Bullet parsing
// ---------------------------------------------------------------------------

/** The bullet font, size and type fields of neutral paragraph properties. */
export function bulletFieldsFromDrawing(
	properties: DrawingParagraphProperties,
): Pick<
	BulletInfo,
	| 'fontFamily'
	| 'fontPanose'
	| 'fontPitchFamily'
	| 'fontCharset'
	| 'fontInherit'
	| 'sizePercent'
	| 'sizePts'
	| 'sizeInherit'
	| 'colorInherit'
> {
	const font = properties.bulletFont;
	const size = properties.bulletSize;
	return {
		...(typeof font === 'object'
			? {
					fontFamily: font.typeface,
					...(font.panose ? { fontPanose: font.panose } : {}),
					...(font.pitchFamily !== undefined ? { fontPitchFamily: font.pitchFamily } : {}),
					...(font.charset !== undefined ? { fontCharset: font.charset } : {}),
				}
			: {}),
		...(font === 'text' ? { fontInherit: true } : {}),
		...(typeof size === 'object' && size.unit === 'percent'
			? { sizePercent: size.value * 100 }
			: {}),
		...(typeof size === 'object' && size.unit === 'points' ? { sizePts: size.value } : {}),
		...(size === 'text' ? { sizeInherit: true } : {}),
		...(properties.bulletColor === 'text' ? { colorInherit: true } : {}),
	};
}

/**
 * Parse bullet info from paragraph-level properties.
 *
 * Handles `a:buNone`, `a:buChar`, `a:buAutoNum`, along with shared
 * properties `a:buFont`, `a:buSzPct`, `a:buSzPts`, `a:buClr`.
 */
export function parseBulletInfo(
	pPr: XmlObject | undefined,
	paragraphIndex: number = 0,
): BulletInfo | null {
	if (!pPr) {
		return null;
	}
	const properties = objectParagraphProperties(pPr);
	const bullet = properties.bullet;

	// Explicit none. `a:buNone` is a marker element, so it is ALWAYS bare in the
	// file; the codec reads its presence, and SmartArt's "no bullet" is an
	// explicit opinion rather than "no opinion" (lost on save before).
	if (bullet?.kind === 'none') {
		return { none: true };
	}

	const { fontFamily, sizePercent, sizePts } = bulletFieldsFromDrawing(properties);

	const buClr = pPr['a:buClr'] as XmlObject | undefined;
	let color: string | undefined;
	if (buClr) {
		const srgb = buClr['a:srgbClr'] as XmlObject | undefined;
		if (srgb?.['@_val']) {
			color = String(srgb['@_val']);
		}
	}

	if (bullet?.kind === 'char') {
		return { char: bullet.char, fontFamily, sizePercent, sizePts, color };
	}

	// Auto-numbered bullet; an `a:buAutoNum` without its required `type` still
	// numbers (PowerPoint's default scheme).
	if (bullet?.kind === 'autoNumber' || pPr['a:buAutoNum']) {
		return {
			autoNumType: bullet?.kind === 'autoNumber' ? bullet.scheme : undefined,
			autoNumStartAt: (bullet?.kind === 'autoNumber' ? bullet.startAt : undefined) ?? 1,
			paragraphIndex,
			fontFamily,
			sizePercent,
			sizePts,
			color,
		};
	}

	// No bullet
	return null;
}
