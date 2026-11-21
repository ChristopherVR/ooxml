/**
 * Text warp / WordArt logic: Vue port of the React
 * `viewer/utils/warp-path-generators.ts` + `text-warp-classifier.ts`.
 *
 * Pure, framework-agnostic helpers that classify an OOXML `prstTxWarp` preset
 * and build the SVG path `d` attribute for a single warped text baseline.
 *
 * A text element signals WordArt via `element.textStyle.textWarpPreset`
 * (`a:bodyPr/a:prstTxWarp/@prst`), with optional `textWarpAdj` / `textWarpAdj2`
 * adjustment values (raw OOXML 1/60000th units). Use {@link hasTextWarp} to
 * detect a warped element and {@link shouldUseSvgWarp} to decide whether the
 * SVG `<textPath>` renderer applies.
 */
import type { PptxElement, PptxTextWarpPreset, TextSegment } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import { isParagraphSeparatorSegment } from './text-segment-paragraph-break';
import { buildWarpPath } from 'ooxml-core/drawingml';
export {
	buildWarpPath,
	cascadeUpPath,
	cascadeDownPath,
	WARP_PATH_GENERATORS,
} from 'ooxml-core/drawingml';
export type { WarpPathGenerator } from 'ooxml-core/drawingml';

// ── Classifier ──────────────────────────────────────────────────────────

/**
 * Rendering-strategy category for a warp preset.
 *  - `path`:     renders along an SVG `<textPath>` (arcs, waves, circles,
 *                slant/fade/cascade, …)
 *  - `envelope`: the text block stretched between a top and a bottom curve
 *                (plain/inflate/deflate/can). Also
 *                renders along an SVG `<textPath>`; kept as its own category
 *                for documentation/grouping only (see {@link ENVELOPE_PRESETS}).
 *  - `simple`:   unused (kept for API stability; no preset classifies here
 *                any more, see {@link SIMPLE_PRESETS}).
 *  - `none`:     no warp (`textNoShape`, unknown)
 *
 * Every classified preset (`path` and `envelope` alike) should be routed to
 * SVG `<textPath>` rendering via {@link shouldUseSvgWarp}; a binding that
 * instead branches on this category to decide path-vs-CSS-transform will
 * wrongly fall back to a flat CSS-transform approximation for the `envelope`
 * family. Use `shouldUseSvgWarp` for that decision, not `classifyTextWarp`.
 */
export type WarpCategory = 'path' | 'envelope' | 'simple' | 'none';

const NONE_PRESETS: ReadonlySet<string> = new Set(['textNoShape']);

/** Presets that render best with SVG textPath along curved/circular paths. */
const PATH_PRESETS: ReadonlySet<string> = new Set([
	'textArchUp',
	'textArchDown',
	'textCircle',
	'textWave1',
	'textWave2',
	'textWave4',
	'textDoubleWave1',
	'textCurveUp',
	'textCurveDown',
	'textArchUpPour',
	'textArchDownPour',
	'textCirclePour',
	'textButton',
	'textButtonPour',
	'textRingInside',
	'textRingOutside',
	'textTriangle',
	'textTriangleInverted',
	'textChevron',
	'textChevronInverted',
	'textStop',
	// Moved from the former `simple` (CSS-transform) family: all six have
	// single-line-safe generators below and render as true SVG textPath,
	// same mechanism as arch/wave/circle.
	'textSlantUp',
	'textSlantDown',
	'textFadeUp',
	'textFadeDown',
	'textFadeLeft',
	'textFadeRight',
	'textCascadeUp',
	'textCascadeDown',
]);

/**
 * Presets that stretch the text block between a top and a bottom curve
 * (plain/inflate/deflate/can). `textPlain`'s curves are the box's straight
 * top and bottom edges, so it stretches the block to fill the shape.
 *
 * Historically these were CSS-transform-approximated (see
 * {@link getEnvelopeCssTransform}), but every one of them already has a
 * single-line-safe SVG baseline generator in {@link WARP_PATH_GENERATORS}
 * (see the "single-line degeneracy" note there), so `classifyTextWarp` keeps
 * classifying them as `'envelope'` for documentation/grouping purposes only;
 * the actual path-vs-CSS rendering decision every binding should make is
 * {@link shouldUseSvgWarp}, which is `true` for all of them.
 */
const ENVELOPE_PRESETS: ReadonlySet<string> = new Set([
	'textPlain',
	'textInflate',
	'textDeflate',
	'textInflateBottom',
	'textInflateTop',
	'textDeflateBottom',
	'textDeflateTop',
	'textDeflateInflate',
	'textDeflateInflateDeflate',
	'textCanUp',
	'textCanDown',
]);

/**
 * Historically "presets approximated by basic 2D transforms (slant, fade,
 * cascade)". As of the WordArt envelope fidelity fix, slant/fade/cascade all
 * have single-line-safe SVG baseline generators (like every other classified
 * preset) and are classified as `'path'` instead; this set is now empty and
 * kept only so `WarpCategory` still has a `'simple'` member for API
 * stability. Nothing currently classifies into it.
 */
const SIMPLE_PRESETS: ReadonlySet<string> = new Set([]);

/**
 * Classify a warp preset into a rendering-strategy category.
 *
 * Returns `'none'` for empty/unknown presets so callers can skip rendering
 * without an explicit allowlist check.
 */
export function classifyTextWarp(preset: string | undefined): WarpCategory {
	if (!preset || NONE_PRESETS.has(preset)) {
		return 'none';
	}
	if (PATH_PRESETS.has(preset)) {
		return 'path';
	}
	if (ENVELOPE_PRESETS.has(preset)) {
		return 'envelope';
	}
	if (SIMPLE_PRESETS.has(preset)) {
		return 'simple';
	}
	return 'none';
}

/** All classified warp presets (excludes `none`). */
export const ALL_CLASSIFIED_PRESETS: ReadonlySet<string> = new Set([
	...NONE_PRESETS,
	...PATH_PRESETS,
	...ENVELOPE_PRESETS,
	...SIMPLE_PRESETS,
]);

// ── Paragraph splitter ──────────────────────────────────────────────────

/** A single warp/WordArt paragraph: the runs that flow along one baseline. */
export interface WarpParagraph {
	/** Text segments belonging to this paragraph (no paragraph-break segments). */
	segments: TextSegment[];
}

/** Minimal element shape the warp paragraph splitter needs. */
export interface WarpTextSource {
	text?: string;
	textSegments?: TextSegment[];
}

/**
 * Optional per-segment transform applied while grouping. Returns the segment to
 * emit for the given input segment; callers use it to substitute field text
 * (slide number / date / footer …). Returning the same segment is a no-op.
 */
export type WarpSegmentTransform = (segment: TextSegment) => TextSegment;

/**
 * Group an element's `textSegments` into paragraphs delimited by paragraph
 * SEPARATOR segments (the break markers themselves are excluded): see
 * {@link isParagraphSeparatorSegment} for what counts as one. A caller that
 * checked only `isParagraphBreak` here used to miss the far more common
 * slide-load case (a bare `text: '\n'` segment with no flag at all), so a
 * freshly-opened multi-paragraph WordArt block never split into separate
 * lines: the literal `"\n"` fell into the first paragraph's own run list and
 * was measured and rendered as its own glyph, which also flipped the
 * vertical ordering of a multi-line envelope (paragraph 2's text ended up
 * sharing paragraph 1's line instead of occupying its own band).
 *
 * Falls back to a single synthetic paragraph carrying `element.text` when no
 * segments are present, and to an empty array when there is neither text nor
 * segments.
 *
 * When `transform` is supplied each non-break segment is passed through it
 * before being collected; this is how the React renderer substitutes field
 * values (slide number, date, footer, …) into warped text. The transform is a
 * separate concern kept out of this pure splitter so Vue/Angular can call it
 * without pulling in field-substitution logic.
 *
 * @param source    Element (or element-like) carrying `text` / `textSegments`.
 * @param transform Optional per-segment substitution callback.
 */
export function groupIntoParagraphs(
	source: WarpTextSource,
	transform?: WarpSegmentTransform,
): WarpParagraph[] {
	const segments = source.textSegments;
	if (!segments || segments.length === 0) {
		if (source.text) {
			return [{ segments: [{ text: source.text, style: {} }] }];
		}
		return [];
	}
	const paragraphs: WarpParagraph[] = [];
	let current: TextSegment[] = [];
	for (const seg of segments) {
		if (isParagraphSeparatorSegment(seg)) {
			if (current.length > 0) {
				paragraphs.push({ segments: current });
			}
			current = [];
		} else {
			current.push(transform ? transform(seg) : seg);
		}
	}
	if (current.length > 0) {
		paragraphs.push({ segments: current });
	}
	return paragraphs;
}

// ── Path generators ─────────────────────────────────────────────────────

/** Presets that require SVG textPath rendering (all others fall back to flat). */
export const SVG_WARP_PRESETS: ReadonlySet<string> = new Set([
	// Priority 1
	'textArchUp',
	'textArchDown',
	'textCircle',
	'textWave1',
	'textInflate',
	'textDeflate',
	'textCurveUp',
	'textCurveDown',
	// Priority 2
	'textWave2',
	'textWave4',
	'textDoubleWave1',
	'textCanUp',
	'textCanDown',
	'textButton',
	'textRingInside',
	'textRingOutside',
	'textCascadeUp',
	'textCascadeDown',
	// Priority 3
	'textTriangle',
	'textTriangleInverted',
	'textStop',
	'textChevron',
	'textChevronInverted',
	'textInflateBottom',
	'textInflateTop',
	'textDeflateBottom',
	'textDeflateTop',
	// Priority 4 – slant, fade, pour, and compound deflate/inflate
	'textSlantUp',
	'textSlantDown',
	'textFadeRight',
	'textFadeLeft',
	'textFadeUp',
	'textFadeDown',
	'textArchUpPour',
	'textArchDownPour',
	'textCirclePour',
	'textButtonPour',
	'textDeflateInflate',
	'textDeflateInflateDeflate',
	'textPlain',
]);

// ── Public API ──────────────────────────────────────────────────────────

/** Returns `true` when the preset should use SVG `<textPath>` rendering. */
export function shouldUseSvgWarp(preset: PptxTextWarpPreset | undefined): boolean {
	if (!preset || preset === 'textNoShape') {
		return false;
	}
	return SVG_WARP_PRESETS.has(preset);
}

/**
 * Predicate: does this element carry a WordArt/text-warp preset that the SVG
 * `<textPath>` renderer should handle? The warp field lives on
 * `element.textStyle.textWarpPreset`.
 */
export function hasTextWarp(element: PptxElement): boolean {
	if (!hasTextProperties(element)) {
		return false;
	}
	return shouldUseSvgWarp(element.textStyle?.textWarpPreset);
}

/**
 * Alias for {@link buildWarpPath}, kept under the original React/Angular symbol
 * name (`getWarpPath`) so binding renderers can import it directly from shared
 * without renaming their call sites.
 */
export const getWarpPath = buildWarpPath;

// ── CSS-transform approximation (envelope / simple) ─────────────────────

/**
 * Framework-agnostic CSS-transform approximation for a warp preset.
 *
 * `path` presets render along an SVG `<textPath>`; `envelope`/`simple`
 * presets cannot bend individual glyphs, so the visual effect is hinted with
 * a `transform` (perspective / rotateX / rotateY / skew / scale) plus the
 * `transform-origin` that anchors it. Port of the React `getEnvelopeCssTransform`
 * + `getSimpleCssTransform` generators (`viewer/utils/text-warp-classifier.ts`)
 * and the simpler `getTextWarpStyle` map (`viewer/utils/text-warp-css.tsx`).
 */
export interface WarpCssTransform {
	/** The CSS `transform` value. */
	transform: string;
	/** The CSS `transform-origin` value. */
	transformOrigin: string;
}

/**
 * Default OOXML adjustment values for envelope presets (raw 1/60000th units).
 * Default `adj1` of 18750 maps to an intensity factor of 1.
 */
const ENVELOPE_ADJ_DEFAULTS: Readonly<Record<string, number>> = {
	textInflate: 18750,
	textDeflate: 18750,
	textInflateBottom: 18750,
	textInflateTop: 18750,
	textDeflateBottom: 18750,
	textDeflateTop: 18750,
	textDeflateInflate: 18750,
	textDeflateInflateDeflate: 18750,
	textCanUp: 18750,
	textCanDown: 18750,
};

/**
 * CSS-transform approximation for an `envelope` preset (inflate/deflate/can).
 *
 * Envelope warps distort text non-uniformly (e.g. inflate widens the middle
 * lines, narrows the top/bottom). CSS cannot bend glyphs, so `scaleX`/`scaleY`,
 * `perspective`, and `rotateX` give a reasonable hint scaled by the `adj1`
 * intensity. Returns `undefined` for a non-envelope preset.
 *
 * @param preset One of the envelope warp preset names.
 * @param adj1   Optional first adjustment value (raw OOXML units).
 * @param adj2   Optional second adjustment value (raw OOXML units; reserved).
 */
export function getEnvelopeCssTransform(
	preset: string,
	adj1?: number,
	adj2?: number,
): WarpCssTransform | undefined {
	const defaultAdj1 = ENVELOPE_ADJ_DEFAULTS[preset] ?? 18750;
	const a1 = adj1 ?? defaultAdj1;
	void adj2;

	// Normalise adj1 to a 0..4 intensity factor (default adj 18750 -> factor 1).
	const intensity = Math.max(0, Math.min(a1 / 18750, 4));

	switch (preset) {
		case 'textInflate':
			return {
				transform: `scaleY(${1 + 0.15 * intensity}) scaleX(${1 + 0.05 * intensity})`,
				transformOrigin: 'center center',
			};
		case 'textInflateBottom':
			return {
				transform: `perspective(${600 - 100 * intensity}px) rotateX(${-8 * intensity}deg)`,
				transformOrigin: 'center bottom',
			};
		case 'textInflateTop':
			return {
				transform: `perspective(${600 - 100 * intensity}px) rotateX(${8 * intensity}deg)`,
				transformOrigin: 'center top',
			};
		case 'textDeflate':
			return {
				transform: `scaleY(${1 - 0.12 * intensity}) scaleX(${1 - 0.05 * intensity})`,
				transformOrigin: 'center center',
			};
		case 'textDeflateBottom':
			return {
				transform: `perspective(${600 - 100 * intensity}px) rotateX(${6 * intensity}deg)`,
				transformOrigin: 'center bottom',
			};
		case 'textDeflateTop':
			return {
				transform: `perspective(${600 - 100 * intensity}px) rotateX(${-6 * intensity}deg)`,
				transformOrigin: 'center top',
			};
		case 'textDeflateInflate':
			return {
				transform: `scaleY(${1 - 0.08 * intensity}) scaleX(${1 + 0.04 * intensity})`,
				transformOrigin: 'center center',
			};
		case 'textDeflateInflateDeflate':
			return {
				transform: `scaleY(${1 - 0.15 * intensity}) scaleX(${1 + 0.06 * intensity})`,
				transformOrigin: 'center center',
			};
		case 'textCanUp':
			return {
				transform: `perspective(${500 - 80 * intensity}px) rotateX(${-6 * intensity}deg)`,
				transformOrigin: 'center center',
			};
		case 'textCanDown':
			return {
				transform: `perspective(${500 - 80 * intensity}px) rotateX(${6 * intensity}deg)`,
				transformOrigin: 'center center',
			};
		default:
			return undefined;
	}
}

/** Default OOXML adjustment values for simple presets (raw 1/60000th units). */
const SIMPLE_ADJ_DEFAULTS: Readonly<Record<string, number>> = {
	textSlantUp: 55000,
	textSlantDown: 55000,
	textFadeRight: 50000,
	textFadeLeft: 50000,
	textFadeUp: 50000,
	textFadeDown: 50000,
	textCascadeUp: 44444,
	textCascadeDown: 44444,
};

/**
 * CSS-transform approximation for a `simple` preset (slant/fade/cascade).
 *
 * These warps are well-modelled by basic 2D transforms (`skewY`, `perspective`,
 * `rotateX`/`rotateY`) scaled by the `adj1` value. Returns `undefined` for a
 * non-simple preset.
 *
 * @param preset One of the simple warp preset names.
 * @param adj1   Optional adjustment value (raw OOXML units).
 */
export function getSimpleCssTransform(preset: string, adj1?: number): WarpCssTransform | undefined {
	const a1 = adj1 ?? SIMPLE_ADJ_DEFAULTS[preset] ?? 50000;

	switch (preset) {
		case 'textSlantUp':
			return {
				transform: `perspective(500px) rotateY(${8 * (a1 / 55000)}deg) skewY(${-4 * (a1 / 55000)}deg)`,
				transformOrigin: 'left center',
			};
		case 'textSlantDown':
			return {
				transform: `perspective(500px) rotateY(${-8 * (a1 / 55000)}deg) skewY(${4 * (a1 / 55000)}deg)`,
				transformOrigin: 'right center',
			};
		case 'textFadeRight':
			return {
				transform: `perspective(400px) rotateY(${-10 * (a1 / 50000)}deg)`,
				transformOrigin: 'left center',
			};
		case 'textFadeLeft':
			return {
				transform: `perspective(400px) rotateY(${10 * (a1 / 50000)}deg)`,
				transformOrigin: 'right center',
			};
		case 'textFadeUp':
			return {
				transform: `perspective(400px) rotateX(${-10 * (a1 / 50000)}deg)`,
				transformOrigin: 'center bottom',
			};
		case 'textFadeDown':
			return {
				transform: `perspective(400px) rotateX(${10 * (a1 / 50000)}deg)`,
				transformOrigin: 'center top',
			};
		case 'textCascadeUp':
			return {
				transform: `skewY(${-8 * (a1 / 44444)}deg)`,
				transformOrigin: 'left top',
			};
		case 'textCascadeDown':
			return {
				transform: `skewY(${8 * (a1 / 44444)}deg)`,
				transformOrigin: 'left top',
			};
		default:
			return undefined;
	}
}

/**
 * Dispatch a preset to its CSS-transform approximation based on its
 * {@link classifyTextWarp} category.
 *
 * Returns the `envelope`/`simple` transform, or `undefined` for `path`/`none`
 * presets (which render via `<textPath>` or flat text respectively). This is
 * the CSS-transform counterpart to {@link buildWarpPath}.
 *
 * @param preset OOXML `prstTxWarp` preset name.
 * @param adj    Optional primary adjustment (raw OOXML 1/60000th units).
 * @param adj2   Optional secondary adjustment (raw OOXML 1/60000th units).
 */
export function getWarpCssTransform(
	preset: PptxTextWarpPreset | string | undefined,
	adj?: number,
	adj2?: number,
): WarpCssTransform | undefined {
	const category = classifyTextWarp(preset);
	if (category === 'envelope') {
		return getEnvelopeCssTransform(preset as string, adj, adj2);
	}
	if (category === 'simple') {
		return getSimpleCssTransform(preset as string, adj);
	}
	return undefined;
}
