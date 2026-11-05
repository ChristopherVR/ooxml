/**
 * Shape-following outer shadow and glow.
 *
 * A rectangle (or an ellipse / round rectangle, whose `border-radius` is part of
 * its own box) paints its outer shadow and glow as `box-shadow`. Every other
 * preset, and every custom geometry, paints its fill as a CSS `clip-path` on the
 * same element, and `clip-path` is applied AFTER `filter` and `box-shadow`, so
 * it removed the shadow and the glow along with the corners of the box. A shape
 * with no fill was worse: its `box-shadow` is the shadow of the invisible box,
 * where PowerPoint shadows only the outline.
 *
 * Such a shape now paints like this:
 *
 * - the container loses its clip-path (the geometry cascade answers `none`),
 * - the fill is painted by the per-sub-path SVG overlay (`subpath-fill-overlay`),
 *   which is clipped by nothing,
 * - the outer shadow and the glow are ONE SVG `<filter>` on the container
 *   (`filter: url(#shape-fx-<id>)`) computed from the painted alpha: fill,
 *   outline and text together, so they follow the silhouette and a hollow shape
 *   shadows its outline only.
 *
 * The profiles are the rectangle's (`effect-profile.ts`): a shadow is a Gaussian
 * of deviation `0.327 * blurRad`; a glow is the silhouette grown by `0.417 * rad`
 * (a ring of offset copies, since `feMorphology` is a square kernel and would
 * grow diagonal edges by 1.41x) and blurred with deviation `rad / 6`.
 *
 * Only shapes whose other paint stays out of the way qualify: the container's
 * `box-shadow` must carry nothing else (no inner shadow, no line shadow, no 3-D
 * bevel) and the fill must be one the overlay can paint (a solid colour or none).
 *
 * @module render/silhouette-effects
 */
import type { PptxElement, ShapeStyle } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import { glowCssLengths, shadowBlurToCss } from './effect-profile';
import { resolveBaseShapeGeometry } from './shape-geometry-base';
import { getSubpathGeometry } from './subpath-geometry';

/** One outer shadow layer, in the units of the box-shadow it replaces. */
export interface SilhouetteShadowLayer {
	readonly offsetX: number;
	readonly offsetY: number;
	/** The CSS blur length (twice the Gaussian's deviation). */
	readonly blur: number;
	/** `#RRGGBB`. */
	readonly color: string;
	readonly opacity: number;
}

/** A glow: the silhouette grown by `spread` then blurred by `blur` (a CSS length). */
export interface SilhouetteGlow {
	readonly spread: number;
	readonly blur: number;
	readonly color: string;
	readonly opacity: number;
}

/** An inner shadow: the offset points to the edge it hugs (as in PowerPoint). */
export interface SilhouetteInnerShadow {
	readonly offsetX: number;
	readonly offsetY: number;
	/** The CSS blur length (twice the Gaussian's deviation). */
	readonly blur: number;
	readonly color: string;
	readonly opacity: number;
}

/** The effects a silhouette filter paints. */
export interface SilhouetteEffectParams {
	readonly shadows: readonly SilhouetteShadowLayer[];
	readonly glow?: SilhouetteGlow;
	readonly inner?: SilhouetteInnerShadow;
}

/** What a binding needs to apply the effects of one shape. */
export interface SilhouetteEffects {
	readonly filterId: string;
	/** The `filter` value that applies it. */
	readonly cssReference: string;
	/** The `<filter>` markup, to inject next to the soft-edge filter. */
	readonly filterMarkup: string;
	/** The filter also paints the inner shadow, so the box-shadow must not. */
	readonly paintsInnerShadow: boolean;
}

const DEFAULT_SHADOW = '#000000';
const DEFAULT_GLOW = '#ffff00';

/** Offsets on the outer ring of the glow dilation, and on the inner ring. */
const GLOW_RING_STEPS = 24;
const GLOW_INNER_RING_STEPS = 12;

function unit(value: number): number {
	return Math.min(1, Math.max(0, value));
}

function hex(value: string | undefined, fallback: string): string {
	if (!value || value === 'transparent') {
		return fallback;
	}
	const candidate = value.startsWith('#') ? value : `#${value}`;
	return /^#[0-9A-Fa-f]{6}$/u.test(candidate) ? candidate : fallback;
}

function finite(value: number | undefined): value is number {
	return typeof value === 'number' && Number.isFinite(value);
}

/** The id of the filter a shape's silhouette effects are painted with. */
export function getSilhouetteFilterId(elementId: string): string {
	return `shape-fx-${elementId.replace(/[^A-Za-z0-9_.-]/gu, '_')}`;
}

/**
 * The outer shadow layers of a style, exactly as the `box-shadow` path resolves
 * them (`getOuterShadowCss` / `getMultiLayerShadowCss`).
 */
export function resolveOuterShadowLayers(
	style: ShapeStyle,
	rotation: number | undefined,
): SilhouetteShadowLayer[] {
	const layers: SilhouetteShadowLayer[] = [];
	if (style.shadows && style.shadows.length > 0) {
		for (const shadow of style.shadows) {
			if (!shadow.color || shadow.color === 'transparent') {
				continue;
			}
			const angle = ((shadow.angle ?? 0) * Math.PI) / 180;
			const distance = shadow.distance ?? 0;
			layers.push({
				offsetX: Math.cos(angle) * distance,
				offsetY: Math.sin(angle) * distance,
				blur: Math.max(0, shadow.blur ?? 6),
				color: hex(shadow.color, DEFAULT_SHADOW),
				opacity: unit(shadow.opacity ?? 0.35),
			});
		}
		return layers;
	}
	if (!style.shadowColor || style.shadowColor === 'transparent') {
		return layers;
	}
	const ooxml = style.outerShadowXml !== undefined;
	let offsetX: number;
	let offsetY: number;
	if (finite(style.shadowAngle) && finite(style.shadowDistance)) {
		const angle =
			style.shadowRotateWithShape === false && typeof rotation === 'number'
				? style.shadowAngle - rotation
				: style.shadowAngle;
		offsetX = Math.cos((angle * Math.PI) / 180) * style.shadowDistance;
		offsetY = Math.sin((angle * Math.PI) / 180) * style.shadowDistance;
	} else {
		offsetX = finite(style.shadowOffsetX) ? style.shadowOffsetX : ooxml ? 0 : 4;
		offsetY = finite(style.shadowOffsetY) ? style.shadowOffsetY : ooxml ? 0 : 4;
	}
	const rawBlur = finite(style.shadowBlur) ? Math.max(0, style.shadowBlur) : ooxml ? 0 : 6;
	layers.push({
		offsetX,
		offsetY,
		blur: ooxml ? shadowBlurToCss(rawBlur) : rawBlur,
		color: hex(style.shadowColor, DEFAULT_SHADOW),
		opacity: finite(style.shadowOpacity) ? unit(style.shadowOpacity) : ooxml ? 1 : 0.35,
	});
	return layers;
}

/** The glow of a style as dilation and blur lengths, or `undefined`. */
export function resolveGlow(style: ShapeStyle): SilhouetteGlow | undefined {
	if (!style.glowColor || style.glowColor === 'transparent' || !style.glowRadius) {
		return undefined;
	}
	if (!(style.glowRadius > 0)) {
		return undefined;
	}
	const { spread, blur } = glowCssLengths(style.glowRadius);
	return {
		spread,
		blur,
		color: hex(style.glowColor, DEFAULT_GLOW),
		opacity: finite(style.glowOpacity) ? unit(style.glowOpacity) : 0.75,
	};
}

/** The inner shadow of a style, with the numbers `getInnerShadowCss` uses. */
export function resolveInnerShadow(style: ShapeStyle): SilhouetteInnerShadow | undefined {
	if (!style.innerShadowColor || style.innerShadowColor === 'transparent') {
		return undefined;
	}
	return {
		offsetX: finite(style.innerShadowOffsetX) ? style.innerShadowOffsetX : 0,
		offsetY: finite(style.innerShadowOffsetY) ? style.innerShadowOffsetY : 0,
		blur: finite(style.innerShadowBlur) ? Math.max(0, style.innerShadowBlur) : 6,
		color: hex(style.innerShadowColor, DEFAULT_SHADOW),
		opacity: finite(style.innerShadowOpacity) ? unit(style.innerShadowOpacity) : 0.5,
	};
}

/** Whether a style paints no fill (the fill overlay would paint nothing either). */
function paintsNoFill(style: ShapeStyle | undefined): boolean {
	return (
		!style || !style.fillColor || style.fillColor === 'transparent' || style.fillMode === 'none'
	);
}

/** Fill modes the SVG overlay can paint (a solid colour, or none). */
function overlayCanPaintFill(style: ShapeStyle | undefined): boolean {
	const mode = style?.fillMode;
	return mode === undefined || mode === 'solid' || mode === 'none';
}

/** Fill modes painted by a clipped child layer (gradient, pattern, picture). */
function isLayerFillMode(style: ShapeStyle | undefined): boolean {
	const mode = style?.fillMode;
	return mode === 'gradient' || mode === 'pattern' || mode === 'image';
}

/** The effect parameters of a shape that paints them as a silhouette, or `undefined`. */
export function getSilhouetteEffectParams(
	element: PptxElement,
): SilhouetteEffectParams | undefined {
	if (element.type !== 'shape' || !hasShapeProperties(element)) {
		return undefined;
	}
	const style = element.shapeStyle;
	if (!style) {
		return undefined;
	}
	const shadows = resolveOuterShadowLayers(style, element.rotation);
	const glow = resolveGlow(style);
	const innerShadow = resolveInnerShadow(style);
	if (shadows.length === 0 && !glow && !innerShadow) {
		return undefined;
	}
	// Anything else that rides on the container's box-shadow would draw across
	// the whole box once the clip is gone.
	if (style.lineShadowColor || style.lineGlowColor || style.shape3d || style.scene3d) {
		return undefined;
	}
	const layerFill = isLayerFillMode(style);
	if (!layerFill && !overlayCanPaintFill(style)) {
		return undefined;
	}
	const geometry = resolveBaseShapeGeometry(element);
	// A gradient, pattern or picture fill moves to a clipped child layer, which
	// only a clipped shape needs; the others keep their box-shadow.
	if (layerFill && geometry.kind !== 'clipPath') {
		return undefined;
	}
	if (geometry.kind === 'bare' || geometry.kind === 'strokeOnly' || geometry.kind === 'lineEdge') {
		return undefined;
	}
	// A filled rectangle, ellipse or round rectangle already has the right
	// `box-shadow`; it is only wrong when the box is not the shape (a clip-path) or
	// when there is no fill to shadow.
	// A translucent fill shows the shadow through it, which a box-shadow (painted
	// outside the box only) cannot.
	const translucent = typeof style.fillOpacity === 'number' && style.fillOpacity < 0.999;
	if (geometry.kind !== 'clipPath' && !paintsNoFill(style) && !translucent) {
		return undefined;
	}
	// The fill of a clipped shape moves to the SVG overlay, which needs geometry.
	if (
		geometry.kind === 'clipPath' &&
		!layerFill &&
		!paintsNoFill(style) &&
		!getSubpathGeometry(element, true)
	) {
		return undefined;
	}
	// An inner shadow rides on the silhouette only where the clip-path removed the
	// box it was painted in (a clipped shape with a fill to shade).
	const inner =
		innerShadow && geometry.kind === 'clipPath' && !paintsNoFill(style) ? innerShadow : undefined;
	// Nothing left that the box-shadow cannot paint.
	if (shadows.length === 0 && !glow && !inner) {
		return undefined;
	}
	return { shadows, ...(glow ? { glow } : {}), ...(inner ? { inner } : {}) };
}

function px(value: number): string {
	return String(Math.round(value * 100) / 100);
}

function glowMarkup(glow: SilhouetteGlow): string {
	const parts: string[] = [];
	const names: string[] = [];
	const ring = (radius: number, steps: number, tag: string): void => {
		for (let i = 0; i < steps; i++) {
			const angle = (2 * Math.PI * i) / steps;
			const name = `${tag}${i}`;
			names.push(name);
			parts.push(
				`<feOffset in="SourceAlpha" dx="${px(Math.cos(angle) * radius)}" dy="${px(Math.sin(angle) * radius)}" result="${name}"/>`,
			);
		}
	};
	if (glow.spread > 0.01) {
		ring(glow.spread, GLOW_RING_STEPS, 'gr');
		ring(glow.spread / 2, GLOW_INNER_RING_STEPS, 'gi');
	}
	const merge = [
		...names.map((n) => `<feMergeNode in="${n}"/>`),
		'<feMergeNode in="SourceAlpha"/>',
	];
	parts.push(`<feMerge result="glowGrown">${merge.join('')}</feMerge>`);
	parts.push(
		`<feGaussianBlur in="glowGrown" stdDeviation="${px(glow.blur / 2)}" result="glowBlur"/>`,
	);
	parts.push(
		`<feFlood flood-color="${glow.color}" flood-opacity="${px(glow.opacity)}" result="glowColor"/>`,
		`<feComposite in="glowColor" in2="glowBlur" operator="in" result="glowLayer"/>`,
	);
	return parts.join('');
}

/** The `<filter>` markup that paints `params` around an element of `width` x `height`. */
export function buildSilhouetteFilterMarkup(
	filterId: string,
	width: number,
	height: number,
	params: SilhouetteEffectParams,
): string {
	let reach = 0;
	for (const s of params.shadows) {
		reach = Math.max(reach, Math.hypot(s.offsetX, s.offsetY) + 1.5 * s.blur);
	}
	if (params.glow) {
		reach = Math.max(reach, params.glow.spread + 1.5 * params.glow.blur);
	}
	if (params.inner) {
		reach = Math.max(
			reach,
			Math.hypot(params.inner.offsetX, params.inner.offsetY) + 1.5 * params.inner.blur,
		);
	}
	const pad = Math.ceil(reach) + 4;
	const parts: string[] = [
		`<filter id="${filterId}" filterUnits="userSpaceOnUse" x="${-pad}" y="${-pad}" width="${px(Math.max(width, 1) + 2 * pad)}" height="${px(Math.max(height, 1) + 2 * pad)}" color-interpolation-filters="sRGB">`,
	];
	const layers: string[] = [];
	if (params.glow) {
		parts.push(glowMarkup(params.glow));
		layers.push('glowLayer');
	}
	params.shadows.forEach((s, i) => {
		const name = `shadow${i}`;
		parts.push(
			`<feGaussianBlur in="SourceAlpha" stdDeviation="${px(s.blur / 2)}" result="${name}Blur"/>`,
			`<feOffset in="${name}Blur" dx="${px(s.offsetX)}" dy="${px(s.offsetY)}" result="${name}Off"/>`,
			`<feFlood flood-color="${s.color}" flood-opacity="${px(s.opacity)}" result="${name}Color"/>`,
			`<feComposite in="${name}Color" in2="${name}Off" operator="in" result="${name}"/>`,
		);
		layers.push(name);
	});
	// Bottom to top: the glow, the shadows (the last one lowest, like a box-shadow
	// list), then the shape itself.
	const stack = [
		...layers.slice(0, params.glow ? 1 : 0),
		...layers.slice(params.glow ? 1 : 0).reverse(),
	];
	const top: string[] = [];
	if (params.inner) {
		const i = params.inner;
		parts.push(
			'<feComponentTransfer in="SourceAlpha" result="innerInv"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer>',
			`<feOffset in="innerInv" dx="${px(-i.offsetX)}" dy="${px(-i.offsetY)}" result="innerOff"/>`,
			`<feGaussianBlur in="innerOff" stdDeviation="${px(i.blur / 2)}" result="innerBlur"/>`,
			`<feFlood flood-color="${i.color}" flood-opacity="${px(i.opacity)}" result="innerFlood"/>`,
			'<feComposite in="innerFlood" in2="innerBlur" operator="in" result="innerColor"/>',
			'<feComposite in="innerColor" in2="SourceAlpha" operator="in" result="innerLayer"/>',
		);
		top.push('innerLayer');
	}
	parts.push(
		`<feMerge>${[...stack, 'SourceGraphic', ...top].map((n) => `<feMergeNode in="${n}"/>`).join('')}</feMerge>`,
		'</filter>',
	);
	return parts.join('');
}

/**
 * The silhouette effects of a shape, or `undefined` when its shadow and glow
 * ride on its `box-shadow` as before.
 */
export function getSilhouetteEffects(element: PptxElement): SilhouetteEffects | undefined {
	const params = getSilhouetteEffectParams(element);
	if (!params) {
		return undefined;
	}
	const filterId = getSilhouetteFilterId(element.id);
	return {
		filterId,
		cssReference: `url(#${filterId})`,
		filterMarkup: buildSilhouetteFilterMarkup(filterId, element.width, element.height, params),
		paintsInnerShadow: params.inner !== undefined,
	};
}

/**
 * Whether the fill of a shape with silhouette effects is painted by the SVG
 * overlay: it was clipped by a clip-path, and has a fill to paint.
 */
export function silhouetteFillsOverlay(element: PptxElement): boolean {
	if (!usesSilhouetteEffects(element) || !hasShapeProperties(element)) {
		return false;
	}
	return (
		resolveBaseShapeGeometry(element).kind === 'clipPath' &&
		!paintsNoFill(element.shapeStyle) &&
		!isLayerFillMode(element.shapeStyle)
	);
}

/**
 * Whether the fill of a shape with silhouette effects is a CSS background on a
 * clipped child layer (`silhouette-fill-layer`): a gradient, pattern or picture
 * fill, which an SVG path cannot paint. The clip applies to the layer before the
 * container's filter, so the filter still sees the silhouette.
 */
export function silhouetteFillsLayer(element: PptxElement): boolean {
	return (
		hasShapeProperties(element) &&
		isLayerFillMode(element.shapeStyle) &&
		usesSilhouetteEffects(element)
	);
}

/** Whether a shape paints its outer shadow and glow as a silhouette. */
export function usesSilhouetteEffects(element: PptxElement): boolean {
	return getSilhouetteEffectParams(element) !== undefined;
}
