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
import { namedFontVerticalMetrics, officeFontVerticalMetrics } from 'ooxml-core/text';

import { typoLineGapEm } from './font-line-gap';
import { firstFontFamily, fontState } from './font-settled';
import type { ParagraphRun, RenderParagraph } from './paragraph-types';
import { type BodyTop, bodyTopSnapShift } from './text-body-top-snap';
import { lineHeightToPx, proportionalLineHeight, resolveLineHeight } from './text-line-height';

/** A font's ascent and descent in px, as CSS lays a line box out with them. */
export interface FontBox {
	ascent: number;
	descent: number;
	/**
	 * The line gap (same unit as `ascent`) PowerPoint adds above the ascent
	 * before scaling the box to its pitch; browsers do not expose it (see
	 * `font-line-gap`). Absent means none.
	 */
	lineGap?: number;
	/**
	 * Set when the box describes a stand-in: the font is a web font that has
	 * not loaded yet (or is missing), so the browser measured a fallback face.
	 * Such a box is never cached, and a paragraph placed from it is settled
	 * again once the fonts have loaded (see `paragraph-baseline-settle`).
	 */
	provisional?: boolean;
	/** Of a provisional box: the font is a web font that is still loading, so its own box is about to be known. */
	loading?: boolean;
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
	const single = size * SINGLE_PITCH;
	if (pitch > single * (1 + SINGLE_EPSILON)) {
		return share;
	}
	// Exactly single spacing keeps the font's own descent below the baseline,
	// however deep (Segoe Script's 0.375 em would otherwise be lifted to 75%);
	// the 75% floor only holds the tighter spacings.
	return pitch >= single * (1 - SINGLE_EPSILON)
		? pitch - descent
		: Math.max(share, pitch - descent);
}

/**
 * The descent (px) PowerPoint keeps below the baseline of a single line of
 * `size` px text: the font's ascent and descent scaled to fill its 1.2 pitch.
 * `box` is the font's ascent and descent at any one size.
 */
export function powerPointDescent(size: number, box: FontBox): number {
	const total = box.ascent + box.descent + (box.lineGap ?? 0);
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

const isDigit = (text: string, index: number): boolean => {
	const code = text.charCodeAt(index);
	return code >= 48 && code <= 57;
};

/**
 * The first `<n>px` size in a CSS `font` shorthand (`NaN` when there is none),
 * the leftmost match of `\d+(\.\d+)?px`. It scans back from each `px` rather
 * than running that pattern, which backtracks quadratically on long digit runs
 * that never reach a `px`.
 */
export function fontSizePx(font: string): number {
	for (let at = font.indexOf('px'); at >= 0; at = font.indexOf('px', at + 2)) {
		let start = at;
		while (start > 0 && isDigit(font, start - 1)) start--;
		if (start === at) continue;
		if (start > 1 && font[start - 1] === '.' && isDigit(font, start - 2)) {
			start--;
			while (start > 0 && isDigit(font, start - 1)) start--;
		}
		return Number.parseFloat(font.slice(start, at));
	}
	return Number.NaN;
}

let context: CanvasRenderingContext2D | null | undefined;
const cache = new Map<string, FontBox>();

/**
 * {@link FontBoxMeasure} on a 2D canvas (`fontBoundingBoxAscent` /
 * `fontBoundingBoxDescent`, the metrics the browser lays a line box out
 * with). Answers `undefined` outside a browser. A font that has not settled
 * (see `isFontSettled`: a web font still loading, or one that is missing)
 * is measured on the stand-in the browser paints it with, flagged
 * `provisional` and left out of the cache, so a stand-in's box is never
 * mistaken for the font's own once it has loaded.
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
	if (!context) {
		return undefined;
	}
	const state = fontState(font);
	context.font = font;
	const metrics = context.measureText('x');
	const box: FontBox = {
		ascent: metrics.fontBoundingBoxAscent,
		descent: metrics.fontBoundingBoxDescent,
	};
	if (!Number.isFinite(box.ascent) || !Number.isFinite(box.descent)) {
		return undefined;
	}
	if (state !== 'settled') {
		return { ...box, provisional: true, ...(state === 'loading' ? { loading: true } : {}) };
	}
	const size = fontSizePx(font);
	const gap = typoLineGapEm(firstFontFamily(font)) * size;
	if (gap > 0) {
		box.lineGap = gap;
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
	/** Where the paragraph's text body starts, for the pixel snap of its top (see `text-body-top-snap`). */
	bodyTop?: BodyTop;
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
 * its lines cannot be modelled: no text, runs in more than one face or (except
 * at an exact pitch above every run's single spacing) more than one size,
 * shifted (superscript, subscript) runs, equations, tabs, ruby or script
 * fallback pieces. Every line of such a paragraph has the same pitch and the
 * same baseline offset, so one move fixes them all and leaves the flow alone.
 */
export function resolveParagraphBaselineShift(input: BaselineShiftInput): number | undefined {
	return resolveParagraphBaseline(input)?.shift;
}

/** What {@link baselineShiftOf} reads: one line's size, pitch and font boxes. */
export interface BaselineGeometry {
	/** Font size (px) of the text and its line pitch (px). */
	size: number;
	pitch: number;
	/** The inline boxes on the line: the paragraph strut, then the text run. */
	boxes: readonly [CssLineBox, CssLineBox];
	/** The text's font box at a large size, whose proportions PowerPoint scales to its pitch. */
	reference: FontBox;
	/** The paragraph strut's font size (px). */
	strutSizePx: number;
	/**
	 * The boxes are those of a stand-in for a web font that is still loading.
	 * Its own proportions are not known yet, so the move is only what the
	 * spacing changes from single spacing, which the renderer leaves where the
	 * browser puts it; the paragraph is placed again when the font has loaded.
	 */
	loading?: boolean;
}

/** The baseline model of PowerPoint's tighter spacings before per-font proportions (see `loading`). */
function standInBaseline(pitch: number, size: number, descent: number): number {
	const share = BASELINE_SHARE * pitch;
	return pitch > size * SINGLE_PITCH * (1 + SINGLE_EPSILON)
		? share
		: Math.max(share, pitch - descent);
}

/** How far (px) to move a line so its baseline is PowerPoint's, or `undefined` for no move. */
export function baselineShiftOf(geometry: BaselineGeometry): number | undefined {
	const { size, pitch, boxes, reference, strutSizePx } = geometry;
	let shift: number;
	if (geometry.loading) {
		const [strut, text] = boxes;
		const single = SINGLE_PITCH * size;
		const singleBoxes = [
			{ ...strut, lineHeight: SINGLE_PITCH * strutSizePx },
			{ ...text, lineHeight: single },
		];
		shift =
			standInBaseline(pitch, size, text.font.descent) -
			standInBaseline(single, size, text.font.descent) -
			(cssBaseline(boxes) - cssBaseline(singleBoxes));
	} else {
		shift =
			powerPointBaseline(pitch, size, powerPointDescent(size, reference)) - cssBaseline(boxes);
	}
	return Math.abs(shift) < MIN_SHIFT_PX ? undefined : Math.round(shift * 100) / 100;
}

/** The font a paragraph's text is measured at for the proportions PowerPoint scales to its pitch. */
export function referenceFont(style: ParagraphRun['style'], family: string): string {
	const italic = style.fontStyle === 'italic' ? 'italic ' : '';
	return `${italic}${String(style.fontWeight ?? 'normal')} ${REFERENCE_SIZE}px ${family}`;
}

/**
 * The box PowerPoint scales to its pitch. A Windows or Office font that is
 * missing here (a legacy East Asian one, an Office cloud font, a styled face
 * name) is laid out by PowerPoint on its own metrics, not on the stand-in the
 * browser measured; a web font still loading keeps its path.
 */
export function referenceBox(
	measured: FontBox | undefined,
	family: string | undefined,
): FontBox | undefined {
	const named =
		measured?.provisional &&
		!measured.loading &&
		(namedFontVerticalMetrics(family) ?? officeFontVerticalMetrics(family));
	return named
		? {
				ascent: named.ascent * REFERENCE_SIZE,
				descent: named.descent * REFERENCE_SIZE,
				...(named.lineGap ? { lineGap: named.lineGap * REFERENCE_SIZE } : {}),
				provisional: true,
			}
		: measured;
}

/** {@link resolveParagraphBaselineShift} that also says whether a stand-in font was measured. */
export function resolveParagraphBaseline(
	input: BaselineShiftInput,
): { shift: number | undefined; provisional: boolean; extentDelta?: number } | undefined {
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
	if (runs.some((run) => !(sizeOf(run) > 0) || familyOf(run) !== family)) {
		return undefined;
	}
	if (runs.some((run) => sizeOf(run) !== size)) {
		return mixedSizeBaseline(
			input,
			runs.map((run) => ({ run, size: sizeOf(run) })),
			family,
		);
	}
	const strut = measure(`normal ${strutSizePx}px ${strutFamily}`);
	const text = measure(runFont(first, size, family));
	// The browser rounds a font's ascent and descent to whole px at the size it
	// lays a line out in, which is right for the CSS baseline but too coarse for
	// the proportions PowerPoint scales to its pitch: those come from a large size.
	const reference = referenceBox(
		measure(referenceFont(first.style, family)) ?? text,
		firstFontFamily(`1px ${family}`),
	);
	if (!strut || !text || !reference) {
		return undefined;
	}
	let pitch: number;
	let boxes: [CssLineBox, CssLineBox];
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
	const provisional = Boolean(strut.provisional || text.provisional || reference.provisional);
	const loading = Boolean(strut.loading || text.loading || reference.loading);
	const extentDelta = loading ? undefined : lastLineExtentDelta(size, pitch, reference);
	const baseShift = baselineShiftOf({ size, pitch, boxes, reference, strutSizePx, loading });
	const snap = input.bodyTop ? bodyTopSnapShift(input.bodyTop, size) : 0;
	const total = Math.round(((baseShift ?? 0) + snap) * 100) / 100;
	return {
		shift: Math.abs(total) < MIN_SHIFT_PX ? undefined : total,
		provisional,
		...(extentDelta === undefined ? {} : { extentDelta }),
	};
}

/**
 * A line whose runs differ in size, at an exact pitch above every run's single
 * spacing. PowerPoint then puts the baseline at {@link BASELINE_SHARE} of the
 * pitch whatever the sizes (a 9pt bullet opening a 7pt line at 11pt exact sits
 * level with an all-7pt line), so only the CSS baseline of the strut and every
 * run's box is needed, and the font's own proportions are not. Any other mix is
 * left where the browser puts it.
 */
function mixedSizeBaseline(
	input: BaselineShiftInput,
	sized: readonly { run: ParagraphRun; size: number }[],
	family: string,
): ReturnType<typeof resolveParagraphBaseline> {
	const { lineHeight, strutSizePx, strutFamily, measure = measureFontBoxWithCanvas } = input;
	if (typeof lineHeight !== 'string' || !/^[\d.]+(?:pt|px)$/u.test(lineHeight.trim())) {
		return undefined;
	}
	const pitch = lineHeightToPx(strutSizePx, lineHeight.trim());
	if (!sized.every(({ size }) => pitch > size * SINGLE_PITCH * (1 + SINGLE_EPSILON))) {
		return undefined;
	}
	const strut = measure(`normal ${strutSizePx}px ${strutFamily}`);
	if (!strut || strut.loading) {
		return undefined;
	}
	const boxes: CssLineBox[] = [{ lineHeight: pitch, font: strut }];
	for (const { run, size } of sized) {
		const font = measure(runFont(run, size, family));
		if (!font || font.loading) {
			return undefined;
		}
		boxes.push({ lineHeight: pxOf(run.style.lineHeight) ?? pitch, font });
	}
	const largest = sized.reduce((a, b) => (b.size > a.size ? b : a));
	const reference = referenceBox(
		measure(referenceFont(largest.run.style, family)),
		firstFontFamily(`1px ${family}`),
	);
	const extentDelta =
		reference && !reference.loading
			? lastLineExtentDelta(largest.size, pitch, reference)
			: undefined;
	const snap = input.bodyTop ? bodyTopSnapShift(input.bodyTop, largest.size) : 0;
	const total = Math.round((BASELINE_SHARE * pitch - cssBaseline(boxes) + snap) * 100) / 100;
	return {
		shift: Math.abs(total) < MIN_SHIFT_PX ? undefined : total,
		provisional: false,
		...(extentDelta === undefined ? {} : { extentDelta }),
	};
}

/**
 * How much taller (px) PowerPoint's last line is than its pitch, or `undefined`
 * when it is as tall. PowerPoint ends a text body at the last line's baseline
 * plus the descent, not at the bottom of the line pitch (measured on table
 * rows of one and two paragraphs, Calibri and Arial at 12 to 40pt: a cell
 * grows to `(lines - 1) p + baseline + D + margins`): a pitch above single
 * spacing is shorter than `p` by the quarter of it under the descent, and one
 * below single spacing is taller, since the 75% baseline floor leaves the
 * descent hanging below the line.
 */
export function lastLineExtentDelta(
	size: number,
	pitch: number,
	reference: FontBox,
): number | undefined {
	const descent = powerPointDescent(size, reference);
	const delta = powerPointBaseline(pitch, size, descent) + descent - pitch;
	return Math.abs(delta) < MIN_SHIFT_PX ? undefined : Math.round(delta * 100) / 100;
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
	bodyTop?: BodyTop,
): void {
	if (
		(body?.textDirection && body.textDirection !== 'horizontal') ||
		body?.compatibleLineSpacing ||
		body?.autoFitMode === 'shrink'
	) {
		return;
	}
	const input: BaselineShiftInput = {
		runs: paragraph.runs,
		lineHeight: paragraph.lineHeight ?? bodyLineHeight(body),
		strutSizePx: paragraph.strutFontSizePx ?? block.fontSizePx,
		strutFamily: block.fontFamily,
		...(measure ? { measure } : {}),
		...(bodyTop ? { bodyTop } : {}),
	};
	const placement = baselinePlacementStyle(input);
	if (!placement) {
		return;
	}
	if (placement.shift !== undefined) {
		paragraph.baselineShiftPx = placement.shift;
	}
	if (placement.provisional) {
		// The font is a stand-in until it loads: say what to place again then.
		paragraph.baselineProvisional = true;
	}
	paragraph.paragraphStyle = { ...paragraph.paragraphStyle, ...placement.style };
}

/** What {@link baselinePlacementStyle} answers: the inline style and what it was made from. */
export interface BaselinePlacement {
	style: Record<string, string | number>;
	shift: number | undefined;
	provisional: boolean;
	/** See {@link lastLineExtentDelta}: how much taller the last line is than its pitch (px). */
	extentDelta?: number;
}

/**
 * The inline style that puts one paragraph block on PowerPoint's baselines (a
 * relative `top`, and the pending marker while its font is a stand-in), or
 * `undefined` when it needs none. Shared by shape text and table cells.
 */
export function baselinePlacementStyle(input: BaselineShiftInput): BaselinePlacement | undefined {
	const resolved = resolveParagraphBaseline(input);
	if (!resolved) {
		return undefined;
	}
	const { shift, provisional, extentDelta } = resolved;
	if (shift === undefined && !provisional && extentDelta === undefined) {
		return undefined;
	}
	const style: Record<string, string | number> = {};
	if (shift !== undefined) {
		style.position = 'relative';
		style.top = `${shift}px`;
	}
	if (provisional) {
		style[BASELINE_PENDING_VAR] = encodePendingBaseline(input);
	}
	return { style, shift, provisional, ...(extentDelta === undefined ? {} : { extentDelta }) };
}

/** The custom property that marks a paragraph placed from a stand-in font's box. */
export const BASELINE_PENDING_VAR = '--pptx-baseline-pending';

/**
 * What a placement needs to be redone once its font has loaded, as a value a
 * style declaration can carry in every binding: the first text run's font and
 * the paragraph's line height and strut, URI-encoded JSON.
 */
export function encodePendingBaseline(input: BaselineShiftInput): string {
	const first = input.runs.find((run) => run.text.length > 0);
	const style: Record<string, unknown> = {};
	for (const key of ['fontSize', 'fontFamily', 'fontStyle', 'fontWeight', 'lineHeight'] as const) {
		const value = first?.style[key];
		if (value !== undefined) {
			style[key] = value;
		}
	}
	return encodeURIComponent(
		JSON.stringify({
			style,
			lineHeight: input.lineHeight,
			strutSizePx: input.strutSizePx,
			strutFamily: input.strutFamily,
			...(input.bodyTop ? { bodyTop: input.bodyTop } : {}),
		}),
	);
}

/** The input {@link encodePendingBaseline} recorded, or `undefined` when the value is not one. */
export function decodePendingBaseline(value: string): BaselineShiftInput | undefined {
	try {
		const data = JSON.parse(decodeURIComponent(value.trim())) as {
			style: ParagraphRun['style'];
			lineHeight: number | string;
			strutSizePx: number;
			strutFamily: string;
			bodyTop?: BodyTop;
		};
		return {
			runs: [{ text: 'x', style: data.style }],
			lineHeight: data.lineHeight,
			strutSizePx: data.strutSizePx,
			strutFamily: data.strutFamily,
			...(data.bodyTop ? { bodyTop: data.bodyTop } : {}),
		};
	} catch {
		return undefined;
	}
}
