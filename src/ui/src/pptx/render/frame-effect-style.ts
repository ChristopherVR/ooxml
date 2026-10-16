/**
 * The container effects of an element that has no shape fill or line of its
 * own but still casts effects: a group (`p:grpSpPr/a:effectLst`, cast by its
 * composite raster) and a chart (the chart area's outer shadow and glow,
 * drawn outside the chart frame). Every binding maps the descriptor onto its
 * element container's style instead of keeping its own copy of the branch.
 *
 * @module render/frame-effect-style
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { getComputedEffectStyle } from './visual-effects';

/** CSS a group or chart container carries for its effects. */
export interface FrameEffectStyle {
	filter?: string;
	boxShadow?: string;
	overflow?: 'visible';
}

/**
 * The container effect style of a group or chart, or `undefined` for any
 * other element (whose effects come with its shape style).
 */
export function getFrameEffectStyle(element: PptxElement): FrameEffectStyle | undefined {
	if (element.type !== 'group' && element.type !== 'chart') {
		return undefined;
	}
	const fx = getComputedEffectStyle(element);
	const style: FrameEffectStyle = {};
	if (fx.filter) {
		style.filter = fx.filter;
	}
	// A group's shadow is a composite `drop-shadow` filter; a chart area's is the frame's box.
	if (fx.boxShadow) {
		style.boxShadow = fx.boxShadow;
	}
	if (fx.overflowVisible) {
		style.overflow = 'visible';
	}
	return style;
}
