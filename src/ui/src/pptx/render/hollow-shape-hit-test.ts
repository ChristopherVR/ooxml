/**
 * `hollow-shape-hit-test`: making an UNFILLED shape click-through.
 *
 * PowerPoint hit-tests an unfilled shape on its OUTLINE and its TEXT only: a
 * bare frame drawn over a chart is a border you can grab, not a pane that
 * swallows every click inside it. The web does not work that way - a
 * `background: transparent` box still hit-tests across its whole border box -
 * so a `<a:noFill/>` rectangle laid over other content stole every click meant
 * for what was underneath it, and those elements could only be selected by
 * moving the frame out of the way (issue #132 deck, slide 5, where a panel
 * frame sat over a bar chart and two text boxes).
 *
 * The fix mirrors `connector-hit-target`, which solves the same problem from the
 * other direction: the container goes `pointer-events: none` and a TRANSPARENT
 * stroke along the shape's own outline opts that outline back in with
 * `pointer-events: stroke`. `pointer-events` inherits, so a descendant
 * re-enabling itself under a `none` ancestor works by construction.
 *
 * Deliberately narrow: only a shape with NO text qualifies. An unfilled TEXT
 * box is also outline-only in PowerPoint, but its text must stay clickable and
 * that is a much larger behavioural change than the reported bug needs.
 *
 * @module render/hollow-shape-hit-test
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { getElementTextContent, hasShapeProperties } from 'ooxml-core/pptx';

import { getResolvedShapeClipPath } from './shape-geometry';
import { usesSilhouetteEffects } from './silhouette-effects';
import { outlinePathData } from './stroke-outline';

/**
 * Narrowest a hollow frame's outline hit band may be, in px.
 *
 * WHY a floor: a hairline frame is one or two px of ink, and a target that thin
 * cannot be hit with a mouse, never mind a finger. Matches the generosity
 * `connector-hit-target` already applies to a line.
 */
export const HOLLOW_HIT_MIN_WIDTH = 10;

/** Geometry and band width for a hollow shape's transparent outline target. */
export interface HollowHitOutline {
	/** SVG path data for the shape's own outline. */
	readonly d: string;
	/** Width of the transparent stroke painted along it (0 for a filled target). */
	readonly strokeWidth: number;
	/** The `fill` of the transparent target path (`transparent` hits its interior). */
	readonly fill: 'none' | 'transparent';
	/** The `stroke` of the target path. */
	readonly stroke: 'none' | 'transparent';
	/**
	 * `stroke` for a frame (outline only); `fill` for a silhouette shape, whose
	 * whole geometry (but not the rest of its box) is the click target.
	 */
	readonly pointerEvents: 'stroke' | 'fill';
}

/**
 * Whether PowerPoint would let a click fall THROUGH this shape's interior.
 *
 * True only for a shape that has no fill of any kind and carries no text, i.e.
 * a pure frame.
 *
 * @param element The element to test.
 * @returns `true` when the interior should not capture pointer events.
 */
export function isHollowShapeElement(element: PptxElement): boolean {
	if (element.type !== 'shape' || !hasShapeProperties(element)) {
		return false;
	}
	const style = element.shapeStyle;
	const filled =
		(style?.fillColor !== undefined && style.fillColor !== 'transparent') ||
		style?.fillMode === 'gradient' ||
		style?.fillMode === 'pattern' ||
		style?.fillMode === 'image' ||
		style?.fillMode === 'group' ||
		Boolean(style?.fillGradient);
	if (filled) {
		return false;
	}
	return getElementTextContent(element).trim() === '';
}

/**
 * Build the transparent outline hit target for a hollow shape.
 *
 * @param element The element to build a target for.
 * @returns The outline path and band width, or `undefined` when the element is
 *   not hollow (in which case its box should hit-test normally).
 */
export function buildHollowHitOutline(element: PptxElement): HollowHitOutline | undefined {
	if (!isHollowShapeElement(element)) {
		return buildSilhouetteHitArea(element);
	}
	const d = outlinePathData(getResolvedShapeClipPath(element), element.width, element.height);
	if (!d) {
		return undefined;
	}
	const strokeWidth = hasShapeProperties(element) ? (element.shapeStyle?.strokeWidth ?? 0) : 0;
	return {
		d,
		strokeWidth: Math.max(strokeWidth * 3, HOLLOW_HIT_MIN_WIDTH),
		fill: 'none',
		stroke: 'transparent',
		pointerEvents: 'stroke',
	};
}

/**
 * Whether a shape's container must let clicks through its box and hit-test only
 * its geometry: a frame (outline only) or a silhouette shape (a shape whose
 * shadow and glow are an SVG filter, so it carries no clip-path any more and its
 * box corners would otherwise capture clicks outside the geometry).
 */
export function passesClicksThroughBox(element: PptxElement): boolean {
	return isHollowShapeElement(element) || buildSilhouetteHitArea(element) !== undefined;
}

/**
 * The geometry of a silhouette shape as a transparent click target. The fill
 * used to be a clip-path, which clipped hit testing to the geometry too.
 */
function buildSilhouetteHitArea(element: PptxElement): HollowHitOutline | undefined {
	// Textless only: a text body under a pointer-events:none container could not
	// take the caret while editing, which is a larger change than this needs.
	if (
		element.type !== 'shape' ||
		getElementTextContent(element).trim() !== '' ||
		!usesSilhouetteEffects(element)
	) {
		return undefined;
	}
	const d = outlinePathData(getResolvedShapeClipPath(element), element.width, element.height);
	return d
		? { d, strokeWidth: 0, fill: 'transparent', stroke: 'none', pointerEvents: 'fill' }
		: undefined;
}
