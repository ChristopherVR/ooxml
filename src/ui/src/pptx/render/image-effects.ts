import { buildArtisticFilterMarkup, buildSimpleArtisticCss } from '../../picture-artistic-filters';
import { pictureLuminanceTransfer } from 'ooxml-core/drawingml';
import type { PptxElement, PptxImageEffects } from 'ooxml-core/pptx';
import { isImageLikeElement } from 'ooxml-core/pptx';

import { hexToRgbUnit } from './color-units';
import { getImageCorrectionsFilterTokens, getImageSharpenFilter } from './image-effect-corrections';
import { buildImageBiLevelTable, buildImageLuminanceTransfer } from './image-effect-filter-values';
import {
	getImageFillOverlayFilter,
	getImageFillOverlayFilterId,
	hasImageFillOverlayEffect,
} from './image-fill-overlay';

/**
 * Image-effects composable  -  Vue port of the React `viewer/utils` image-effect
 * layer (`shape-visual-effects.ts`, `duotone-effects.ts`,
 * `shape-visual-filters.tsx`, `artistic-effects.tsx`).
 *
 * It maps the parsed OOXML {@link PptxImageEffects} on a `picture`/`image`
 * element to:
 *   - a CSS `filter:` string (brightness/contrast/saturate/grayscale/blur and
 *     CSS-only artistic approximations), and
 *   - any inline SVG `<filter>` definitions (duotone, advanced alpha primitives,
 *     complex artistic effects) referenced from that CSS string via `url(#id)`.
 *
 * Every function here is pure and unit-testable without mounting a component.
 * The integrator binds the returned CSS `filter` onto the `<img>` and injects
 * the SVG `<filter>` markup once per element (see module docs at bottom).
 *
 * Precedence, unit conversions, and filter IDs intentionally mirror the React
 * implementation so the two bindings render identically.
 */

// ── Public types ────────────────────────────────────────────────────────────

/** An inline SVG `<filter>` definition to inject so a `url(#id)` reference resolves. */
export interface ImageSvgFilterDefinition {
	/** The `id` attribute of the `<filter>` (matches the `url(#id)` reference). */
	id: string;
	/**
	 * Inner markup of the `<filter>` element (the `<fe*>` primitives only).
	 * Wrap it in `<svg width="0" height="0"><defs><filter id=…>…</filter></defs></svg>`.
	 */
	markup: string;
}

/** Aggregate result of {@link getComputedImageStyle}. */
export interface ComputedImageStyle {
	/** CSS `filter` value, or `undefined` when the element has no image effects. */
	filter?: string;
	/** Overall opacity (0–1) derived from `alphaModFix`, or `undefined`. */
	opacity?: number;
	/** SVG `<filter>` defs that must be injected for the CSS `url(#…)` refs to resolve. */
	svgFilters: ImageSvgFilterDefinition[];
}

// ── Internal helpers ────────────────────────────────────────────────────────

/** Clamp a number to the inclusive `[lo, hi]` range. */
function clamp(v: number, lo: number, hi: number): number {
	return v < lo ? lo : v > hi ? hi : v;
}

/** Normalize a radius value (0–100) to a 0–1 float for proportional scaling. */

/** Get the image effects off an element, or `undefined` for non-image elements. */
function getEffects(element: PptxElement): PptxImageEffects | undefined {
	if (!isImageLikeElement(element)) {
		return undefined;
	}
	return element.imageEffects;
}

// ── Filter ID generators (mirror the React util IDs) ────────────────────────

/** Stable SVG filter ID for a duotone effect on an element. */
export function getImageDuotoneFilterId(elementId: string): string {
	return `duotone-${elementId}`;
}

/** Stable SVG filter ID for the advanced-alpha / colour primitives filter. */
export function getImageAlphaFilterId(elementId: string): string {
	return `imgalpha-${elementId}`;
}

/** Stable SVG filter ID for a complex artistic effect on an element. */
export function getArtisticFilterId(elementId: string): string {
	return `artistic-fx-${elementId}`;
}

// ── Artistic effects requiring SVG filters ──────────────────────────────────

/**
 * Set of artistic effect names that require an SVG `<filter>` definition.
 * All others are approximated with CSS-only filter functions.
 */
const SVG_FILTER_EFFECTS = new Set<string>([
	'artisticFilmGrain',
	'filmGrain',
	'artisticCutout',
	'cutout',
	'artisticCement',
	'cement',
	'artisticTexturizer',
	'texturizer',
	'artisticCrisscrossEtching',
	'crisscrossEtching',
	'artisticMosaic',
	'artisticMosaicBubbles',
	// The DrawingML element really is spelled "Mosiaic"; Microsoft's typo is part
	// of the format, so a file-sourced effect only ever arrives under this name.
	'artisticMosiaicBubbles',
	'mosaicBubbles',
	'mosaic',
	'artisticGlowEdges',
	'glowEdges',
	'glow_edges',
	'artisticChalkSketch',
	'chalkSketch',
	'chalk',
	'artisticPencilSketch',
	'pencilSketch',
	'artisticPencilGrayscale',
	'pencilGrayscale',
	'grayPencil',
]);

/**
 * Whether an element's artistic effect has to be drawn by us.
 *
 * PowerPoint applies an artistic effect DESTRUCTIVELY: the bitmap the picture
 * points at already carries the effect, and the `a14` blip extension we parse it
 * from only records the settings (next to the pristine original, kept as a
 * `.wdp` HD Photo part) so the effect can be re-edited. Verified against
 * PowerPoint COM: exporting a slide with and without that extension produces a
 * byte-identical render. Core therefore records the baked effect's name in
 * `artisticPrerenderedEffect`, and only an effect that differs from it (i.e. one
 * chosen in this library's own inspector) is rendered here.
 */
export function isArtisticEffectRendered(
	effects: PptxImageEffects | undefined,
): effects is PptxImageEffects & { artisticEffect: string } {
	return (
		Boolean(effects?.artisticEffect) &&
		effects?.artisticEffect !== effects?.artisticPrerenderedEffect
	);
}

/** Whether an artistic effect name needs an inline SVG `<filter>` (vs pure CSS). */
export function needsSvgArtisticFilter(effectName: string | undefined): boolean {
	if (!effectName) {
		return false;
	}
	return SVG_FILTER_EFFECTS.has(effectName);
}

// ── Advanced-alpha detection ────────────────────────────────────────────────

/**
 * Returns true when the element has any blip-side alpha primitive or advanced
 * colour effect that CSS filters can't express (alphaInv, alphaCeiling,
 * alphaFloor, alphaRepl, alphaBiLevel, biLevel, lum, hsl with sat, tint,
 * clrRepl, alphaMod). Brightness/contrast/saturation/grayscale/duotone are
 * still handled via CSS filters in {@link getImageFilterCss}.
 *
 * `alphaModFix` is deliberately NOT part of this predicate even though it is
 * an alpha primitive: it always has an exact, simpler CSS representation
 * (plain `opacity`, computed by {@link getImageEffectsOpacity}), so it never
 * needs to route through the `imgalpha-<id>` SVG filter this predicate gates.
 * CSS `opacity` is applied by the browser as a final compositing step AFTER
 * any `filter`, so it multiplies correctly with whatever this predicate's
 * other primitives already do to the image's alpha channel; folding
 * alphaModFix into the SVG filter as well would apply it a second time (see
 * `getComputedImageStyle`'s "applies alphaModFix exactly once" tests, and
 * GitHub issue #286, where a 15% alphaModFix rendered at ~2% because both the
 * filter and `opacity` carried the same 0.15 multiplier).
 */
export function hasAdvancedImageAlphaEffects(element: PptxElement): boolean {
	const e = getEffects(element);
	if (!e) {
		return false;
	}
	return Boolean(
		e.alphaInv ||
		e.alphaCeiling ||
		e.alphaFloor ||
		typeof e.alphaMod?.amt === 'number' ||
		typeof e.alphaRepl === 'number' ||
		typeof e.alphaBiLevel === 'number' ||
		typeof e.biLevel === 'number' ||
		(e.lum && (typeof e.lum.bright === 'number' || typeof e.lum.contrast === 'number')) ||
		(e.hsl && (typeof e.hsl.sat === 'number' || typeof e.hsl.lum === 'number')) ||
		(e.tint && typeof e.tint.amt === 'number') ||
		e.clrRepl,
	);
}

// ── CSS filter string ───────────────────────────────────────────────────────

/**
 * Build the CSS `filter:` string for an element's image effects.
 *
 * Mirrors React's `getImageEffectsFilter`: brightness/contrast/saturate are
 * OOXML hundredths-of-percent → CSS multipliers; grayscale → `grayscale(100%)`;
 * duotone / advanced-alpha / complex-artistic effects append a `url(#id)`
 * reference to an SVG `<filter>` that {@link getImageSvgFilters} produces.
 *
 * @param element The slide element (must be `picture`/`image` to yield a value).
 * @param options `excludeDuotone` skips the duotone `url(#…)` ref (e.g. when a
 *                canvas path handles duotone instead).
 * @returns The CSS filter value, or `undefined` when there are no effects.
 */
export function getImageFilterCss(
	element: PptxElement,
	options?: { excludeDuotone?: boolean },
): string | undefined {
	const effects = getEffects(element);
	if (!effects) {
		return undefined;
	}

	const filters: string[] = [];

	// Brightness: OOXML hundredths-of-percent → CSS multiplier
	if (typeof effects.brightness === 'number' && effects.brightness !== 0) {
		filters.push(`brightness(${Math.max(0, 1 + effects.brightness / 100)})`);
	}
	// Contrast: OOXML hundredths-of-percent → CSS multiplier
	if (typeof effects.contrast === 'number' && effects.contrast !== 0) {
		filters.push(`contrast(${Math.max(0, 1 + effects.contrast / 100)})`);
	}
	// Saturation: -100..100 → CSS saturate() multiplier
	if (typeof effects.saturation === 'number' && effects.saturation !== 0) {
		filters.push(`saturate(${Math.max(0, 1 + effects.saturation / 100)})`);
	}
	// Grayscale
	if (effects.grayscale) {
		filters.push('grayscale(100%)');
	}
	// PowerPoint 2010+ Corrections/Color panel (a14 extensions): distinct from,
	// and additive with, the legacy a:blip bright/contrast/saturation above.
	// See image-effect-corrections.ts for the precision and approximation notes.
	filters.push(...getImageCorrectionsFilterTokens(effects));
	const sharpenFilter = getImageSharpenFilter(effects, element.id);
	if (sharpenFilter) {
		filters.push(sharpenFilter.cssReference);
	}
	// Duotone: reference inline SVG filter (rendered by getImageSvgFilters).
	if (effects.duotone && !options?.excludeDuotone) {
		filters.push(`url(#${getImageDuotoneFilterId(element.id)})`);
	}
	// Advanced alpha primitives + biLevel/lum/hsl(sat,lum)/tint/clrRepl that
	// CSS filters can't express; rendered by getImageSvgFilters().
	if (hasAdvancedImageAlphaEffects(element)) {
		filters.push(`url(#${getImageAlphaFilterId(element.id)})`);
	}
	// Blip fill-overlay (a:fillOverlay): flood + blend + clip-to-alpha, built by
	// getImageSvgFilters() via image-fill-overlay.ts.
	if (hasImageFillOverlayEffect(element)) {
		filters.push(`url(#${getImageFillOverlayFilterId(element.id)})`);
	}
	// Artistic effects: complex ones reference the SVG filter; simple ones use
	// CSS filter functions directly. A pre-rendered effect (one PowerPoint baked
	// into the stored bitmap and recorded in the a14 blip extension) is skipped:
	// re-applying it here would double it up. See `isArtisticEffectRendered`.
	if (isArtisticEffectRendered(effects)) {
		const radius = effects.artisticRadius ?? 5;

		if (needsSvgArtisticFilter(effects.artisticEffect)) {
			filters.push(`url(#${getArtisticFilterId(element.id)})`);
		} else {
			filters.push(buildSimpleArtisticCss(effects.artisticEffect, radius));
		}
	}

	return filters.length > 0 ? filters.join(' ') : undefined;
}

/**
 * Legacy alias of {@link getImageFilterCss}, preserved for the React
 * `shape-visual-effects` shim whose public symbol was `getImageEffectsFilter`.
 */
export const getImageEffectsFilter = getImageFilterCss;

/** CSS-only approximations for the simpler artistic effects. */

// ── Opacity ─────────────────────────────────────────────────────────────────

/**
 * Extract overall image opacity from the `alphaModFix` effect.
 * Returns a 0–1 value for CSS `opacity`, or `undefined` if not set.
 *
 * This is the ONLY place `alphaModFix` is applied: `getImageFilterCss` /
 * `hasAdvancedImageAlphaEffects` deliberately never fold it into the
 * `imgalpha-<id>` SVG filter (even when that filter is built for OTHER
 * primitives, e.g. `biLevel`), so combining `alphaModFix` with any other
 * advanced alpha effect still applies the alpha multiply exactly once: once
 * as this CSS `opacity`, composing on top of whatever the filter did.
 */
export function getImageEffectsOpacity(element: PptxElement): number | undefined {
	const effects = getEffects(element);
	if (!effects) {
		return undefined;
	}
	if (typeof effects.alphaModFix === 'number') {
		return clamp(effects.alphaModFix / 100, 0, 1);
	}
	return undefined;
}

// ── Duotone SVG filter ──────────────────────────────────────────────────────

/** BT.601 luminance feColorMatrix values (RGB → grayscale luminance). */
const GRAYSCALE_LUMINANCE_MATRIX = [
	0.2126, 0.7152, 0.0722, 0, 0, 0.2126, 0.7152, 0.0722, 0, 0, 0.2126, 0.7152, 0.0722, 0, 0, 0, 0, 0,
	1, 0,
].join(' ');

/**
 * Build the inner `<filter>` markup for a duotone colour mapping: convert to
 * grayscale luminance, then linearly remap 0→color1 (shadows) and 1→color2
 * (highlights) per channel. Mirrors React's `renderDuotoneSvgFilter`.
 */
function buildDuotoneFilterMarkup(color1: string, color2: string): string {
	const c1 = hexToRgbUnit(color1);
	const c2 = hexToRgbUnit(color2);
	return (
		`<feColorMatrix type="matrix" values="${GRAYSCALE_LUMINANCE_MATRIX}"/>` +
		'<feComponentTransfer>' +
		`<feFuncR type="linear" slope="${c2.r - c1.r}" intercept="${c1.r}"/>` +
		`<feFuncG type="linear" slope="${c2.g - c1.g}" intercept="${c1.g}"/>` +
		`<feFuncB type="linear" slope="${c2.b - c1.b}" intercept="${c1.b}"/>` +
		'</feComponentTransfer>'
	);
}

/**
 * High-fidelity duotone SVG filter for an image element.
 *
 * @returns The filter `id`, the matching `cssReference` (`url(#id)`), and the
 *          inner `filterMarkup` to inject  -  or `undefined` when the element has
 *          no duotone effect.
 */
export function getDuotoneImageFilter(
	element: PptxElement,
	elementId: string = isImageLikeElement(element) ? element.id : '',
): { id: string; cssReference: string; filterMarkup: string } | undefined {
	const effects = getEffects(element);
	if (!effects?.duotone) {
		return undefined;
	}
	const id = getImageDuotoneFilterId(elementId);
	return {
		id,
		cssReference: `url(#${id})`,
		filterMarkup: buildDuotoneFilterMarkup(effects.duotone.color1, effects.duotone.color2),
	};
}

// ── Advanced alpha / colour primitives SVG filter ───────────────────────────

/**
 * Build the inner `<filter>` markup covering the blip-side alpha primitives and
 * advanced colour effects that CSS filters can't express. Mirrors React's
 * `renderImageAlphaSvgFilter` (chained `in`/`result` primitives).
 *
 * @returns The markup, or `undefined` when no such primitives are present.
 */
function buildImageAlphaFilterMarkup(effects: PptxImageEffects): string | undefined {
	const parts: string[] = [];
	let resultIdx = 0;
	let inputRef = 'SourceGraphic';
	const next = (build: (input: string, output: string) => string): void => {
		const output = `r${resultIdx++}`;
		parts.push(build(inputRef, output));
		inputRef = output;
	};

	// Note: `effects.alphaModFix` intentionally has NO primitive here. It is
	// applied exclusively as CSS `opacity` (see `getImageEffectsOpacity` and
	// `hasAdvancedImageAlphaEffects`'s doc comment) so it is never multiplied
	// twice alongside whatever other primitives this filter builds below.

	// a:alphaMod: multiplicative alpha from the nested a:alphaModFix inside its
	// a:cont container (see PptxImageEffects.alphaMod.amt). Distinct from the
	// sibling alphaModFix effect (handled via CSS opacity, not here).
	if (typeof effects.alphaMod?.amt === 'number') {
		const mul = clamp(effects.alphaMod.amt / 100, 0, 1);
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				`<feFuncA type="linear" slope="${mul}" intercept="0"/>` +
				'</feComponentTransfer>',
		);
	}

	if (effects.alphaInv) {
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				'<feFuncA type="linear" slope="-1" intercept="1"/>' +
				'</feComponentTransfer>',
		);
	}

	if (effects.alphaCeiling) {
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				'<feFuncA type="discrete" tableValues="0 1 1 1 1 1 1 1 1 1"/>' +
				'</feComponentTransfer>',
		);
	}

	if (effects.alphaFloor) {
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				'<feFuncA type="discrete" tableValues="0 0 0 0 0 0 0 0 0 1"/>' +
				'</feComponentTransfer>',
		);
	}

	if (typeof effects.alphaRepl === 'number') {
		const a = clamp(effects.alphaRepl / 100, 0, 1);
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				`<feFuncA type="linear" slope="0" intercept="${a}"/>` +
				'</feComponentTransfer>',
		);
	}

	if (typeof effects.alphaBiLevel === 'number') {
		const t = clamp(effects.alphaBiLevel / 100, 0, 1);
		const table = Array.from({ length: 10 }, (_, i) => (i / 10 >= t ? '1' : '0')).join(' ');
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				`<feFuncA type="discrete" tableValues="${table}"/>` +
				'</feComponentTransfer>',
		);
	}

	if (typeof effects.biLevel === 'number') {
		// grayscale, then threshold each channel.
		next((inp, out) => `<feColorMatrix in="${inp}" result="${out}" type="saturate" values="0"/>`);
		const tbl = buildImageBiLevelTable(effects.biLevel);
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				`<feFuncR type="discrete" tableValues="${tbl}"/>` +
				`<feFuncG type="discrete" tableValues="${tbl}"/>` +
				`<feFuncB type="discrete" tableValues="${tbl}"/>` +
				'</feComponentTransfer>',
		);
	}

	if (
		effects.lum &&
		(typeof effects.lum.bright === 'number' || typeof effects.lum.contrast === 'number')
	) {
		const { slope, intercept } = pictureLuminanceTransfer(
			(effects.lum.bright ?? 0) / 100,
			(effects.lum.contrast ?? 0) / 100,
		);
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				`<feFuncR type="linear" slope="${slope}" intercept="${intercept}"/>` +
				`<feFuncG type="linear" slope="${slope}" intercept="${intercept}"/>` +
				`<feFuncB type="linear" slope="${slope}" intercept="${intercept}"/>` +
				'</feComponentTransfer>',
		);
	}

	if (effects.hsl && typeof effects.hsl.sat === 'number') {
		const v = clamp(1 + effects.hsl.sat / 100, 0, 2);
		next(
			(inp, out) => `<feColorMatrix in="${inp}" result="${out}" type="saturate" values="${v}"/>`,
		);
	}

	if (effects.hsl && typeof effects.hsl.lum === 'number') {
		const { slope, intercept } = buildImageLuminanceTransfer(effects.hsl.lum);
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				`<feFuncR type="linear" slope="${slope}" intercept="${intercept}"/>` +
				`<feFuncG type="linear" slope="${slope}" intercept="${intercept}"/>` +
				`<feFuncB type="linear" slope="${slope}" intercept="${intercept}"/>` +
				'</feComponentTransfer>',
		);
	}

	if (effects.tint && typeof effects.tint.hue === 'number') {
		next(
			(inp, out) =>
				`<feColorMatrix in="${inp}" result="${out}" type="hueRotate" values="${effects.tint!.hue}"/>`,
		);
	}

	if (effects.tint && typeof effects.tint.amt === 'number') {
		const { slope, intercept } = buildImageLuminanceTransfer(effects.tint.amt);
		next(
			(inp, out) =>
				`<feComponentTransfer in="${inp}" result="${out}">` +
				`<feFuncR type="linear" slope="${slope}" intercept="${intercept}"/>` +
				`<feFuncG type="linear" slope="${slope}" intercept="${intercept}"/>` +
				`<feFuncB type="linear" slope="${slope}" intercept="${intercept}"/>` +
				'</feComponentTransfer>',
		);
	}

	if (effects.clrRepl) {
		const c = hexToRgbUnit(effects.clrRepl.color);
		next(
			(inp, out) =>
				`<feColorMatrix in="${inp}" result="${out}" type="matrix" ` +
				`values="0 0 0 0 ${c.r}  0 0 0 0 ${c.g}  0 0 0 0 ${c.b}  0 0 0 1 0"/>`,
		);
	}

	return parts.length > 0 ? parts.join('') : undefined;
}

/**
 * Advanced alpha / colour primitives SVG filter for an image element.
 *
 * @returns The filter `id`, its `cssReference`, and the inner `filterMarkup`,
 *          or `undefined` when no advanced primitives apply.
 */
export function getImageAlphaFilter(
	element: PptxElement,
	elementId: string = isImageLikeElement(element) ? element.id : '',
): { id: string; cssReference: string; filterMarkup: string } | undefined {
	const effects = getEffects(element);
	if (!effects || !hasAdvancedImageAlphaEffects(element)) {
		return undefined;
	}
	const markup = buildImageAlphaFilterMarkup(effects);
	if (!markup) {
		return undefined;
	}
	const id = getImageAlphaFilterId(elementId);
	return { id, cssReference: `url(#${id})`, filterMarkup: markup };
}

// ── Artistic SVG filter ─────────────────────────────────────────────────────

/**
 * Complex artistic SVG filter for an image element.
 *
 * @returns The filter `id`, its `cssReference`, and the inner `filterMarkup`,
 *          or `undefined` when the element's artistic effect (if any) is handled
 *          via CSS rather than an SVG filter.
 */
export function getArtisticImageFilter(
	element: PptxElement,
	elementId: string = isImageLikeElement(element) ? element.id : '',
): { id: string; cssReference: string; filterMarkup: string } | undefined {
	const effects = getEffects(element);
	if (!isArtisticEffectRendered(effects) || !needsSvgArtisticFilter(effects.artisticEffect)) {
		return undefined;
	}
	const markup = buildArtisticFilterMarkup(effects.artisticEffect, effects.artisticRadius ?? 5);
	if (!markup) {
		return undefined;
	}
	const id = getArtisticFilterId(elementId);
	return { id, cssReference: `url(#${id})`, filterMarkup: markup };
}

// ── Aggregate ───────────────────────────────────────────────────────────────

/**
 * Collect every SVG `<filter>` definition an element's image effects require
 * (duotone, advanced alpha, complex artistic) in the same order their `url(#…)`
 * references are appended in {@link getImageFilterCss}.
 */
export function getImageSvgFilters(element: PptxElement): ImageSvgFilterDefinition[] {
	const defs: ImageSvgFilterDefinition[] = [];
	const duotone = getDuotoneImageFilter(element);
	if (duotone) {
		defs.push({ id: duotone.id, markup: duotone.filterMarkup });
	}
	const alpha = getImageAlphaFilter(element);
	if (alpha) {
		defs.push({ id: alpha.id, markup: alpha.filterMarkup });
	}
	const artistic = getArtisticImageFilter(element);
	if (artistic) {
		defs.push({ id: artistic.id, markup: artistic.filterMarkup });
	}
	const fillOverlay = getImageFillOverlayFilter(element);
	if (fillOverlay) {
		defs.push({ id: fillOverlay.id, markup: fillOverlay.filterMarkup });
	}
	const effects = getEffects(element);
	const sharpen = effects && getImageSharpenFilter(effects, element.id);
	if (sharpen) {
		defs.push({ id: sharpen.id, markup: sharpen.filterMarkup });
	}
	return defs;
}

/**
 * Aggregate computed style for an image element's effects: the CSS `filter`
 * string, an optional `opacity`, and the SVG `<filter>` defs to inject so the
 * `url(#…)` references in `filter` resolve.
 *
 * Returns empty/undefined fields for non-image elements or elements with no
 * effects, so it is always safe to spread onto an `<img>`.
 */
export function getComputedImageStyle(element: PptxElement): ComputedImageStyle {
	return {
		filter: getImageFilterCss(element),
		opacity: getImageEffectsOpacity(element),
		svgFilters: getImageSvgFilters(element),
	};
}
