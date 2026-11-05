/**
 * Where PowerPoint puts a line's baseline inside its line pitch, and how far
 * CSS has to move a paragraph to match.
 *
 * CSS centres the line height on the font's own ascent and descent (half the
 * leading above, half below), so changing `line-height` by some amount moves
 * every baseline by half of it. PowerPoint does not. Measured on PowerPoint 16
 * slide exports (one-line paragraphs, 49 font and size pairs at single spacing
 * and four fonts at `a:spcPct` 70% to 200% and `a:spcPts` 30 to 60pt, read at
 * 4px per point), with `p` the line pitch, `s` the font size and `D` the
 * descent PowerPoint keeps below the baseline:
 *
 * - Single spacing is a pitch of exactly `1.2 s` for every font, with the
 *   font's own ascent and descent (the Windows metrics, `winAscent` and
 *   `winDescent`) scaled to fill it: the baseline sits `1.2 s a / (a + d)`
 *   from the top of the line, so `D = 1.2 s d / (a + d)`. Calibri measured
 *   0.940 em, Arial 0.975, Georgia 0.969, Segoe UI 0.975, Verdana 0.994.
 * - A line taller than single spacing (`p > 1.2 s`) sits its baseline at
 *   exactly 75% of the pitch, whatever the font: 102% to 200% and every exact
 *   pitch above single all measured `0.75 p` (plus the glyph overshoot).
 * - At single spacing and below, the descent stays at the bottom of the line
 *   and the space is taken above it, `p - D` (100%, 95% and 90%), until that
 *   falls under `0.75 p` (85% and tighter), which holds again.
 *
 * So `baseline = p > 1.2 s ? 0.75 p : max(0.75 p, p - D)`. The jump at single
 * spacing is real: 102% put the Calibri baseline 0.75pt HIGHER than 100% did.
 *
 * The move is the whole difference between that baseline and the one CSS
 * gives the line (see {@link cssBaseline}), single spacing included: fonts
 * differ in where CSS puts their single line (Arial and Georgia sat 1.5pt
 * high at 40pt, Segoe UI 1.4pt low), so no font is left to the flow.
 *
 * @module render/paragraph-baseline
 */
import type { TextStyle } from 'ooxml-core/pptx';

import type { ParagraphRun, RenderParagraph } from './paragraph-types';
import { lineHeightToPx, proportionalLineHeight, resolveLineHeight } from './text-line-height';

/** A font's ascent and descent in px, as CSS lays a line box out with them. */
export interface FontBox {
	ascent: number;
	descent: number;
}

/** Ascent and descent of a CSS `font` shorthand, or `undefined` when they cannot be measured. */
export type FontBoxMeasure = (font: string) => FontBox | undefined;

/** PowerPoint's single line pitch as a multiple of the font size. */
const SINGLE_PITCH = 1.2;
/** Where PowerPoint puts the baseline in a line that is not at or under single spacing. */
const BASELINE_SHARE = 0.75;
/** A pitch this close to single spacing counts as single. */
const SINGLE_EPSILON = 0.001;
/** The font size (px) the ascent and descent proportions are measured at. */
const REFERENCE_SIZE = 1000;
/** Moves smaller than this (px) are left alone. */
const MIN_SHIFT_PX = 0.05;

/** PowerPoint's baseline, in px from the top of a line `pitch` tall of text `size` px. */
export function powerPointBaseline(pitch: number, size: number, descent: number): number {
	const share = BASELINE_SHARE * pitch;
	return pitch > size * SINGLE_PITCH * (1 + SINGLE_EPSILON)
		? share
		: Math.max(share, pitch - descent);
}

/**
 * The descent (px) PowerPoint keeps below the baseline of a single line of
 * `size` px text: the font's ascent and descent scaled to fill its 1.2 pitch.
 * `box` is the font's ascent and descent at any one size.
 */
export function powerPointDescent(size: number, box: FontBox): number {
	const total = box.ascent + box.descent;
	return total > 0 ? (SINGLE_PITCH * size * box.descent) / total : box.descent;
}

/** One inline box CSS puts on a line: its `line-height` in px and its font's box. */
export interface CssLineBox {
	lineHeight: number;
	font: FontBox;
}

/**
 * The baseline CSS gives a line holding `boxes`, in px from the line's top.
 * Layout works on whole-pixel ascents and descents and floors the leading it
 * puts above a box (`ascent + floor((lineHeight - ascent - descent) / 2)`),
 * which moves a line by up to half a pixel against the unrounded centring.
 */
export function cssBaseline(boxes: readonly CssLineBox[]): number {
	return Math.max(
		...boxes.map(({ lineHeight, font }) => {
			const ascent = Math.round(font.ascent);
			const descent = Math.round(font.descent);
			return ascent + Math.floor((lineHeight - ascent - descent) / 2);
		}),
	);
}

let context: CanvasRenderingContext2D | null | undefined;
const cache = new Map<string, FontBox>();

/**
 * {@link FontBoxMeasure} on a 2D canvas (`fontBoundingBoxAscent` /
 * `fontBoundingBoxDescent`, the metrics the browser lays a line box out
 * with). Answers `undefined` outside a browser and for a web font that has
 * not loaded yet, and caches only fonts that are ready.
 */
export const measureFontBoxWithCanvas: FontBoxMeasure = (font) => {
	const known = cache.get(font);
	if (known) {
		return known;
	}
	if (context === undefined) {
		context =
			typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
	}
	const fonts = typeof document === 'undefined' ? undefined : document.fonts;
	if (!context || (fonts && !fonts.check(font))) {
		return undefined;
	}
	context.font = font;
	const metrics = context.measureText('x');
	const box = { ascent: metrics.fontBoundingBoxAscent, descent: metrics.fontBoundingBoxDescent };
	if (!Number.isFinite(box.ascent) || !Number.isFinite(box.descent)) {
		return undefined;
	}
	cache.set(font, box);
	return box;
};

/** A CSS length in px, from a number or a `"<n>px"` string. */
function pxOf(value: unknown): number | undefined {
	if (typeof value === 'number') {
		return value;
	}
	const parsed = typeof value === 'string' ? Number.parseFloat(value) : Number.NaN;
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** What {@link resolveParagraphBaselineShift} reads from one paragraph. */
export interface BaselineShiftInput {
	runs: readonly ParagraphRun[];
	/** The paragraph's `line-height`: its own, else the text body's. */
	lineHeight: number | string;
	/** The paragraph element's own `font-size` (its strut) and `font-family`. */
	strutSizePx: number;
	strutFamily: string;
	measure?: FontBoxMeasure;
}

/** The canvas font of a run at `size` in `family`. */
function runFont(run: ParagraphRun, size: number, family: string): string {
	const italic = run.style.fontStyle === 'italic' ? 'italic ' : '';
	return `${italic}${String(run.style.fontWeight ?? 'normal')} ${size}px ${family}`;
}

/** Whether a run lays out as plain text on the paragraph's baseline. */
function plainRun(run: ParagraphRun): boolean {
	return (
		!run.equation &&
		!run.tabLines &&
		!run.ruby &&
		!run.scriptRuns &&
		(run.style.verticalAlign === undefined || run.style.verticalAlign === 'baseline')
	);
}

/**
 * How far (px) to move a paragraph down so its baselines sit where PowerPoint
 * puts them (see the module comment), or `undefined` when it needs no move or
 * its lines cannot be modelled: no text, runs in more than one size or face,
 * shifted (superscript, subscript) runs, equations, tabs, ruby or script
 * fallback pieces. Every line of such a paragraph has the same pitch and the
 * same baseline offset, so one move fixes them all and leaves the flow alone.
 */
export function resolveParagraphBaselineShift(input: BaselineShiftInput): number | undefined {
	const { lineHeight, strutSizePx, strutFamily, measure = measureFontBoxWithCanvas } = input;
	const runs = input.runs.filter((run) => run.text.length > 0);
	const first = runs[0];
	if (!first || !runs.every(plainRun)) {
		return undefined;
	}
	const sizeOf = (run: ParagraphRun) => pxOf(run.style.fontSize) ?? strutSizePx;
	const familyOf = (run: ParagraphRun) => ownFamily(run.style.fontFamily) ?? strutFamily;
	const size = sizeOf(first);
	const family = familyOf(first);
	if (!(size > 0) || runs.some((run) => sizeOf(run) !== size || familyOf(run) !== family)) {
		return undefined;
	}
	const strut = measure(`normal ${strutSizePx}px ${strutFamily}`);
	const text = measure(runFont(first, size, family));
	// The browser rounds a font's ascent and descent to whole px at the size it
	// lays a line out in, which is right for the CSS baseline but too coarse for
	// the proportions PowerPoint scales to its pitch: those come from a large size.
	const referenceBox = measure(runFont(first, REFERENCE_SIZE, family));
	if (!strut || !text) {
		return undefined;
	}
	let pitch: number;
	let boxes: CssLineBox[];
	if (typeof lineHeight === 'number') {
		pitch = lineHeight * size;
		boxes = [
			{ lineHeight: lineHeight * strutSizePx, font: strut },
			{ lineHeight: pitch, font: text },
		];
	} else {
		pitch = lineHeightToPx(size, lineHeight);
		boxes = [
			{ lineHeight: pitch, font: strut },
			{ lineHeight: pxOf(first.style.lineHeight) ?? pitch, font: text },
		];
	}
	const shift =
		powerPointBaseline(pitch, size, powerPointDescent(size, referenceBox ?? text)) -
		cssBaseline(boxes);
	return Math.abs(shift) < MIN_SHIFT_PX ? undefined : Math.round(shift * 100) / 100;
}

/** A run's own `font-family`, when it sets one. */
function ownFamily(value: unknown): string | undefined {
	return typeof value === 'string' && value.length > 0 ? value : undefined;
}

/** The `line-height` a text body gives paragraphs that set none, as `buildTextBlockStyle` does. */
function bodyLineHeight(style: TextStyle | undefined): number | string {
	const reduction = style?.autoFitLineSpacingReduction;
	if (style?.autoFitMode !== 'shrink' && typeof reduction === 'number' && reduction > 0) {
		return (
			proportionalLineHeight(style?.lineSpacing, style?.compatibleLineSpacing) * (1 - reduction)
		);
	}
	return resolveLineHeight(style, false);
}

/**
 * Move `paragraph` onto PowerPoint's baselines (see
 * {@link resolveParagraphBaselineShift}) with a relative `top`, which leaves
 * every other paragraph where it is. Skipped for vertical text, the legacy
 * `compatLnSpc` model and a box the browser shrinks to fit at run time.
 */
export function applyParagraphBaselineShift(
	paragraph: RenderParagraph,
	body: TextStyle | undefined,
	block: { fontFamily: string; fontSizePx: number },
	measure?: FontBoxMeasure,
): void {
	if (
		(body?.textDirection && body.textDirection !== 'horizontal') ||
		body?.compatibleLineSpacing ||
		body?.autoFitMode === 'shrink'
	) {
		return;
	}
	const shift = resolveParagraphBaselineShift({
		runs: paragraph.runs,
		lineHeight: paragraph.lineHeight ?? bodyLineHeight(body),
		strutSizePx: paragraph.strutFontSizePx ?? block.fontSizePx,
		strutFamily: block.fontFamily,
		...(measure ? { measure } : {}),
	});
	if (shift === undefined) {
		return;
	}
	paragraph.baselineShiftPx = shift;
	paragraph.paragraphStyle = {
		...paragraph.paragraphStyle,
		position: 'relative',
		top: `${shift}px`,
	};
}
