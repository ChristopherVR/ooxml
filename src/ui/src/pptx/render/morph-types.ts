/**
 * Types, interfaces, and constants for the PowerPoint Morph transition system.
 *
 * Pure, framework-agnostic. Consumed by every binding; bindings keep only the
 * DOM/RAF animation driver.
 *
 * @module render/morph-types
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { easingSegment } from './morph-easing';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** A matched pair of elements between the outgoing and incoming slides. */
export interface MorphPair {
	fromElement: PptxElement;
	toElement: PptxElement;
}

/** Result from full morph matching including unmatched elements. */
export interface MorphMatchResult {
	/** Matched element pairs to animate between. */
	pairs: MorphPair[];
	/** Element IDs only present on the outgoing (from) slide  -  these fade out. */
	unmatchedFrom: PptxElement[];
	/** Element IDs only present on the incoming (to) slide  -  these fade in. */
	unmatchedTo: PptxElement[];
}

/** Describes the CSS animation and keyframes for a single morph-animated element. */
export interface MorphAnimationStyle {
	elementId: string;
	/** CSS animation string. */
	animation: string;
	/** Inline keyframes block to inject. */
	keyframes: string;
	/**
	 * Which node inside the element the animation belongs on.
	 *
	 * `element` (the default) is the element's own positioned container, the
	 * node that carries its static transform. `image` is the `<img>` INSIDE
	 * that container, which every binding renders for a picture: a source crop
	 * (`a:srcRect`) is painted by transforming the img within its frame, so
	 * morphing a crop change has to animate that node and not the frame - the
	 * frame is usually exactly the same size on both slides (issue #148).
	 * `outline` is the `<path>` of the outline overlay every binding stamps
	 * `data-pptx-outline` on, which a resizing shape re-paths per frame (see
	 * `morph-shape-carry`).
	 */
	target?: MorphAnimationTarget;
}

/** @see MorphAnimationStyle.target */
export type MorphAnimationTarget = 'element' | 'image' | 'outline';

/** Morph granularity mode matching PowerPoint's morph effect options. */
export type MorphMode = 'object' | 'word' | 'character';

/** A single token (word or character) with its computed position for text morphing. */
export interface MorphTextToken {
	text: string;
	/** Normalised x offset within the text frame (0-1). */
	x: number;
	/** Normalised y offset within the text frame (0-1). */
	y: number;
	fontSize: number;
	fontWeight: string;
	color: string;
}

/** Paired tokens for text morph animation. */
export interface MorphTextTokenPair {
	from: MorphTextToken | null;
	to: MorphTextToken | null;
}

/** Parsed RGBA colour for interpolation. */
export interface RgbaColor {
	r: number;
	g: number;
	b: number;
	a: number;
}

/** A single SVG path command with its coordinate values. */
export interface SvgPathCommand {
	type: string;
	values: number[];
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/**
 * Control points of PowerPoint's morph easing curve, `[x1, y1, x2, y2]`.
 * Anything that samples the curve itself (the stepped crop track, the dissolve
 * windows below) reads these, so the CSS string and the sampled keyframes
 * cannot drift apart.
 *
 * Measured on PowerPoint 16 for issue #141. Synthetic 960x540pt decks were
 * built through COM (solid rectangles on white, Morph by object) and exported
 * with `CreateVideo` at 60fps, at morph durations of 1s and 2s. On the one
 * matched pair deck, a rectangle travelling 840pt, another travelling 80pt, a
 * 40x40 box growing to 300x140, a 300x20 bar turning 0 to 90 degrees, a fill
 * turning blue to yellow and a fill going 0% to 80% transparent ALL follow the
 * same normalised progress curve, which does not depend on distance or
 * duration. It reads 0.022 at 5% of the duration, 0.075 at 10%, 0.236 at 20%,
 * 0.419 at 30%, 0.588 at 40%, 0.726 at 50%, 0.832 at 60%, 0.911 at 70%, 0.960
 * at 80% and 0.988 at 90%. This curve tracks the geometry of both durations to
 * an RMS of 0.0015 (max 0.004) and the colour and transparency to an RMS of
 * 0.003 (max 0.008). The `cubic-bezier(0.4, 0, 0.2, 1)` used before was up to
 * 0.10 behind at the start and 0.05 ahead in the middle.
 */
export const MORPH_EASING_POINTS: readonly [number, number, number, number] = [0.21, 0.02, 0.39, 1];

/** PowerPoint's morph transition uses a specific cubic-bezier easing. */
export const MORPH_EASING = `cubic-bezier(${MORPH_EASING_POINTS.join(', ')})`;

/**
 * The curve a matched pair DISSOLVES on: the same curve it travels on.
 *
 * PowerPoint blends a matched shape's fill colour and transparency on exactly
 * the progress that moves it (see {@link MORPH_EASING_POINTS}), and our ghost
 * crossfade over an opaque incoming half shows that blend, so the dissolve
 * rides the travel curve. The earlier separate fit, `cubic-bezier(0.2, 0, 0.4,
 * 1)` from the issue #131 deck's crossfading title (alpha 0.035 at 7%, 0.232 at
 * 20%, 0.477 at 34%, 0.684 at 47%, 0.888 at 68%, 0.988 at 88%), agrees with
 * this curve to within 0.017 at every one of those samples, so it was the same
 * curve measured on fewer frames. Kept as its own export because the dissolve
 * rides its own animation.
 */
export const MORPH_CROSSFADE_EASING = MORPH_EASING;

/**
 * The dissolve windows of unmatched shapes, in EASED progress (0..1) of
 * {@link MORPH_EASING_POINTS}, not in time.
 *
 * A shape with no counterpart is not cross-faded over the whole transition: it
 * is gone a third of the way along the morph's progress, and its replacement
 * only starts to appear two thirds of the way along, so the middle shows
 * neither. Within each window the alpha is LINEAR in the morph's eased
 * progress. Measured for issue #141 with the same `CreateVideo` method as the
 * curve, alpha read from the luma of the shape's centre (the RGB decode clips
 * saturated fills, which is what made an earlier fit look shape dependent): on
 * fourteen fades (rectangles, ovals and a triangle; leaving, arriving, alone,
 * beside a travelling shape and with both at once; 1s and 2s), fitting the
 * window per fade gives 0.002 to 0.336 for leaving shapes and 0.665 to 1.003
 * for arriving ones, all fourteen within 0.002 of each other. The fixed
 * windows below fit every one of them to an RMS under 0.007. The best window
 * that is linear in TIME instead is 3.5 to 6 times worse.
 */
export const MORPH_FADE_OUT_END_PROGRESS = 1 / 3;
/** @see MORPH_FADE_OUT_END_PROGRESS */
export const MORPH_FADE_IN_START_PROGRESS = 2 / 3;

/** Percent of the duration, to two decimals, for a keyframe selector. */
function toPercent(time: number): number {
	return Math.round(time * 10000) / 100;
}

const FADE_OUT_SEGMENT = easingSegment(MORPH_EASING_POINTS, 0, MORPH_FADE_OUT_END_PROGRESS);
const FADE_IN_SEGMENT = easingSegment(MORPH_EASING_POINTS, MORPH_FADE_IN_START_PROGRESS, 1);

/**
 * When an unmatched OUTGOING shape starts dissolving, as a percentage of the
 * morph's duration. It starts at once: the slow start of the curve is what
 * makes it look as if it holds for the first few frames.
 */
export const MORPH_FADE_OUT_HOLD_PERCENT = 0;
/**
 * When an unmatched OUTGOING shape has finished dissolving, as a percentage of
 * the duration: where the eased progress reaches
 * {@link MORPH_FADE_OUT_END_PROGRESS} (about 25%).
 */
export const MORPH_FADE_OUT_END_PERCENT = toPercent(FADE_OUT_SEGMENT.endTime);
/**
 * The timing function of the leaving dissolve, from 0% to
 * {@link MORPH_FADE_OUT_END_PERCENT}: the slice of the morph curve over that
 * window, so the alpha stays linear in the morph's eased progress.
 */
export const MORPH_FADE_OUT_EASING = FADE_OUT_SEGMENT.easing;

/**
 * When an unmatched INCOMING shape starts dissolving in, as a percentage of the
 * duration: where the eased progress reaches
 * {@link MORPH_FADE_IN_START_PROGRESS} (about 45%). Nothing of it shows before.
 */
export const MORPH_FADE_IN_START_PERCENT = toPercent(FADE_IN_SEGMENT.startTime);
/**
 * The timing function of the arriving dissolve, from
 * {@link MORPH_FADE_IN_START_PERCENT} to 100%: the slice of the morph curve
 * over that window, which decelerates hard into the end.
 */
export const MORPH_FADE_IN_EASING = FADE_IN_SEGMENT.easing;

/** Maximum pixel distance for proximity-based element matching. */
export const PROXIMITY_THRESHOLD = 300;

/**
 * Maximum per-axis size ratio (larger/smaller) for a proximity match. Pairs
 * further apart than this in width OR height dissolve in place instead of
 * gliding, because interpolating the box would visibly stretch one shape into
 * the other (issue #131: a centre-text group paired with a whole wheel wedge).
 */
export const PROXIMITY_SIZE_RATIO_LIMIT = 2;
