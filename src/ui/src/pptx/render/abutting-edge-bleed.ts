/**
 * Closing the hairline seam between two abutting filled rectangles.
 *
 * The stage is one `transform: scale()` over slide-pixel boxes, so a shared
 * edge that sits on a whole slide pixel can land inside a device pixel. The
 * browser anti-aliases each box on its own and composites the two partial
 * coverages, so the background shows through that pixel: a light 1px column
 * between two dark panels (issue #141, slide 13). PowerPoint draws the slide
 * as one surface and shows no seam. Moving the boxes to the device grid does
 * not help, because the browser snaps a box's paint to whole LOCAL pixels
 * before the stage scale applies.
 *
 * The fix is local to such a pair: the LOWER box grows by
 * {@link ABUTTING_EDGE_BLEED_PX} under the UPPER one, so the boundary pixel is
 * covered by one primitive of the lower box and the upper box's partial
 * coverage blends over it. The grown strip is only ever painted where an
 * opaque, unrotated rectangle above it hides it, and the lower box's own fill
 * keeps its authored size (see `abutting-edge-bleed-style.ts`), so nothing
 * that is visible changes except the seam.
 */
import type { PptxElement, PptxSlide, ShapeStyle } from 'ooxml-core/pptx';

/** How far (slide px) a lower box grows under an abutting upper box. */
export const ABUTTING_EDGE_BLEED_PX = 2;

export type AbuttingEdgeSide = 'left' | 'right' | 'top' | 'bottom';

/**
 * One grown side. `color` is the lower fill's colour along that edge when it
 * is uniform there (a solid fill, or a linear gradient running across the
 * edge); absent when the fill runs along the edge, in which case the fill
 * itself extends into the strip.
 */
export interface AbuttingEdgeBleedSide {
	readonly color?: string;
}

/** The sides of one element that grow under an abutting upper box. */
export type AbuttingEdgeBleed = Readonly<Partial<Record<AbuttingEdgeSide, AbuttingEdgeBleedSide>>>;

export interface AbuttingEdgeBleedOptions {
	/** Elements whose box must stay exact (selected or being edited). */
	readonly excludeIds?: ReadonlySet<string>;
	/** Elements that animate: a bleed must never be uncovered mid-effect. */
	readonly animatedIds?: ReadonlySet<string>;
}

const OPAQUE = 0.999;
/** Shared-edge tolerance in slide px (boxes are whole pixels in practice). */
const EDGE_EPSILON = 0.01;

/** Style keys whose presence means the box paints beyond (or less than) its fill. */
const EFFECT_KEY =
	/^(?:shadow|shadows|glow|softEdge|reflection|innerShadow|blur|dag|effectDagTree|scene3d|shape3d|shapeFillOverlay|fillOverlay|lineGlow|lineShadow|useBackgroundFill)/u;

function hasPaintEffects(style: ShapeStyle): boolean {
	for (const [key, value] of Object.entries(style)) {
		if (value !== undefined && value !== false && EFFECT_KEY.test(key)) {
			return true;
		}
	}
	return false;
}

function hasVisibleLine(style: ShapeStyle): boolean {
	return (
		(style.strokeWidth ?? 0) > 0 &&
		style.strokeFillMode !== 'none' &&
		style.strokeColor !== undefined &&
		style.strokeColor !== 'transparent' &&
		(style.strokeOpacity ?? 1) > 0
	);
}

function isPlainRectBox(el: PptxElement): el is PptxElement & { shapeStyle: ShapeStyle } {
	if (el.type !== 'shape' && el.type !== 'text') return false;
	if (el.hidden || el.rotation || el.skewX || el.skewY) return false;
	if (el.width <= 0 || el.height <= 0) return false;
	if ((el.shapeType ?? 'rect').toLowerCase() !== 'rect') return false;
	if ('pathData' in el && (el.pathData || el.customGeometryPaths?.length)) return false;
	const style = el.shapeStyle;
	return style !== undefined && !hasPaintEffects(style);
}

const stopOpacity = (stop: { opacity?: number }): number => stop.opacity ?? 1;

const isPaint = (color: string | undefined): color is string =>
	color !== undefined && color !== '' && color !== 'transparent';

/** An upper box that hides whatever lies under it, everywhere in its box. */
function isOpaqueCover(el: PptxElement): boolean {
	if (!isPlainRectBox(el) || (el.opacity ?? 1) < OPAQUE) return false;
	const style = el.shapeStyle;
	if (style.fillMode === 'solid') {
		return isPaint(style.fillColor) && (style.fillOpacity ?? 1) >= OPAQUE;
	}
	if (style.fillMode === 'gradient') {
		const stops = style.fillGradientStops ?? [];
		return stops.length > 0 && stops.every((stop) => stopOpacity(stop) >= OPAQUE);
	}
	return false;
}

function hasVisibleText(el: PptxElement): boolean {
	if ('placeholderType' in el && el.placeholderType !== undefined) return true;
	if (!('text' in el)) return false;
	if ((el.text ?? '').trim() !== '') return true;
	return (el.textSegments ?? []).some((segment) => (segment.text ?? '').trim() !== '');
}

const isAxisAngle = (angle: number, target: number): boolean =>
	Math.abs((((angle - target) % 360) + 360) % 360) < 0.01 ||
	Math.abs(((((angle - target) % 360) + 360) % 360) - 360) < 0.01;

/**
 * The lower fill's paint along `side`, or `undefined` when the box cannot
 * grow on that side without changing what is visible.
 */
function sidePaint(
	el: PptxElement & { shapeStyle: ShapeStyle },
	side: AbuttingEdgeSide,
): AbuttingEdgeBleedSide | undefined {
	const style = el.shapeStyle;
	if (style.fillMode === 'solid') {
		return isPaint(style.fillColor) && (style.fillOpacity ?? 1) >= OPAQUE
			? { color: style.fillColor }
			: undefined;
	}
	if (style.fillMode !== 'gradient' || style.fillGradientType !== 'linear') return undefined;
	if ((style.fillGradientFlip ?? 'none') !== 'none' || style.fillGradientTileRect) return undefined;
	if (el.flipHorizontal || el.flipVertical) return undefined;
	const stops = [...(style.fillGradientStops ?? [])].sort((a, b) => a.position - b.position);
	const first = stops[0];
	const last = stops[stops.length - 1];
	if (!first || !last) return undefined;
	const angle = style.fillGradientAngle ?? 0;
	const horizontalSide = side === 'left' || side === 'right';
	const across = horizontalSide
		? isAxisAngle(angle, 0) || isAxisAngle(angle, 180)
		: isAxisAngle(angle, 90) || isAxisAngle(angle, 270);
	const along = horizontalSide
		? isAxisAngle(angle, 90) || isAxisAngle(angle, 270)
		: isAxisAngle(angle, 0) || isAxisAngle(angle, 180);
	if (along) return {};
	if (!across) return undefined;
	const startSide = isAxisAngle(angle, 0)
		? 'left'
		: isAxisAngle(angle, 180)
			? 'right'
			: isAxisAngle(angle, 90)
				? 'top'
				: 'bottom';
	const stop = side === startSide ? first : last;
	return stopOpacity(stop) >= OPAQUE ? { color: stop.color } : undefined;
}

interface Box {
	readonly left: number;
	readonly top: number;
	readonly right: number;
	readonly bottom: number;
}

const boxOf = (el: PptxElement): Box => ({
	left: el.x,
	top: el.y,
	right: el.x + el.width,
	bottom: el.y + el.height,
});

const near = (a: number, b: number): boolean => Math.abs(a - b) <= EDGE_EPSILON;

/** The side of `lower` that `upper` abuts across a shared edge, if any. */
function sharedSide(lower: Box, upper: Box): AbuttingEdgeSide | undefined {
	const overlapsY = upper.top < lower.bottom && upper.bottom > lower.top;
	const overlapsX = upper.left < lower.right && upper.right > lower.left;
	if (overlapsY && near(lower.left, upper.right)) return 'left';
	if (overlapsY && near(lower.right, upper.left)) return 'right';
	if (overlapsX && near(lower.top, upper.bottom)) return 'top';
	if (overlapsX && near(lower.bottom, upper.top)) return 'bottom';
	return undefined;
}

/** The strip a grown side paints, corners included (they grow with every side). */
function stripOf(box: Box, side: AbuttingEdgeSide, sides: ReadonlySet<AbuttingEdgeSide>): Box {
	const b = ABUTTING_EDGE_BLEED_PX;
	const left = box.left - (sides.has('left') ? b : 0);
	const right = box.right + (sides.has('right') ? b : 0);
	const top = box.top - (sides.has('top') ? b : 0);
	const bottom = box.bottom + (sides.has('bottom') ? b : 0);
	if (side === 'left') return { left, right: box.left, top, bottom };
	if (side === 'right') return { left: box.right, right, top, bottom };
	if (side === 'top') return { left, right, top, bottom: box.top };
	return { left, right, top: box.bottom, bottom };
}

const contains = (outer: Box, inner: Box): boolean =>
	outer.left <= inner.left + EDGE_EPSILON &&
	outer.right >= inner.right - EDGE_EPSILON &&
	outer.top <= inner.top + EDGE_EPSILON &&
	outer.bottom >= inner.bottom - EDGE_EPSILON;

/**
 * The elements (by id) that grow under an abutting upper box, for one list
 * painted in order (later paints on top). Empty for nearly every slide.
 */
export function resolveAbuttingEdgeBleeds(
	elements: readonly PptxElement[],
	options: AbuttingEdgeBleedOptions = {},
): ReadonlyMap<string, AbuttingEdgeBleed> {
	const result = new Map<string, AbuttingEdgeBleed>();
	const skip = (el: PptxElement): boolean =>
		options.excludeIds?.has(el.id) === true ||
		options.animatedIds?.has(el.id) === true ||
		(el.shapeId !== undefined && options.animatedIds?.has(el.shapeId) === true);
	const covers = elements.map((el) => (!skip(el) && isOpaqueCover(el) ? boxOf(el) : undefined));
	elements.forEach((lower, index) => {
		if (
			skip(lower) ||
			!isPlainRectBox(lower) ||
			hasVisibleLine(lower.shapeStyle) ||
			hasVisibleText(lower)
		) {
			return;
		}
		const box = boxOf(lower);
		const candidates = new Map<AbuttingEdgeSide, { cover: Box; paint: AbuttingEdgeBleedSide }>();
		for (let above = index + 1; above < elements.length; above++) {
			const cover = covers[above];
			const side = cover ? sharedSide(box, cover) : undefined;
			if (!cover || !side || candidates.has(side)) continue;
			const paint = sidePaint(lower, side);
			if (paint) candidates.set(side, { cover, paint });
		}
		// A grown side is kept only while its strip, corners included, lies
		// wholly under its cover; dropping a side only shrinks the others.
		let changed = true;
		while (changed && candidates.size > 0) {
			changed = false;
			const sides = new Set(candidates.keys());
			for (const [side, { cover }] of candidates) {
				if (!contains(cover, stripOf(box, side, sides))) {
					candidates.delete(side);
					changed = true;
					break;
				}
			}
		}
		if (candidates.size > 0) {
			const bleed: Partial<Record<AbuttingEdgeSide, AbuttingEdgeBleedSide>> = {};
			for (const [side, { paint }] of candidates) bleed[side] = paint;
			result.set(lower.id, bleed);
		}
	});
	return result;
}

const NO_BLEEDS: ReadonlyMap<string, AbuttingEdgeBleed> = new Map();

/**
 * The bleeds of one slide stage: its template (layout/master) elements painted
 * first, then the slide's own, as every binding stacks them. Elements the
 * slide animates and the excluded ones (the selection) keep their exact box,
 * and a slide entered by Morph keeps every box exact.
 */
export function resolveStageEdgeBleeds(
	templateElements: readonly PptxElement[] | undefined,
	slide: PptxSlide | undefined,
	excludeIds?: ReadonlySet<string>,
): ReadonlyMap<string, AbuttingEdgeBleed> {
	// A Morph transition paints this stage's elements as separately animated
	// copies, which could move the cover off a grown strip mid-transition.
	if (slide?.transition?.type === 'morph') return NO_BLEEDS;
	const elements = [...(templateElements ?? []), ...(slide?.elements ?? [])];
	if (elements.length < 2) return NO_BLEEDS;
	const bleeds = resolveAbuttingEdgeBleeds(elements, {
		...(excludeIds ? { excludeIds } : {}),
		animatedIds: slideAnimatedElementIds(slide),
	});
	return bleeds.size > 0 ? bleeds : NO_BLEEDS;
}

/** Ids (element ids and shape ids) that a slide's animations target. */
export function slideAnimatedElementIds(slide: PptxSlide | undefined): ReadonlySet<string> {
	const ids = new Set<string>();
	for (const animation of slide?.animations ?? []) ids.add(animation.elementId);
	for (const animation of slide?.nativeAnimations ?? []) {
		if (animation.targetId !== undefined) ids.add(animation.targetId);
	}
	return ids;
}
