/**
 * picture-frame.ts: where a picture's own frame (`p:spPr/a:ln`) and outer
 * shadow (`a:effectLst/a:outerShdw`) are painted, and the box split that lets
 * them show.
 *
 * PowerPoint centres a picture's line on the frame edge, so half of it falls
 * outside the picture, and the outer shadow falls outside it as well. The box
 * that clips the picture's pixels (the `a:srcRect` crop, the picture geometry,
 * the Crop to Shape clip) would cut both. When either is present the clip
 * therefore moves to an inner surface layer, and the box stays unclipped and
 * carries the shadow; the outline overlay (`buildStrokeOutline`) is mounted on
 * the box, beside the surface, and so is the reflection mirror. A picture with
 * neither keeps its clip on the box, so it paints exactly as before.
 *
 * React already paints pictures this way (an unclipped container with the
 * composite shadow, a clipped image surface, and `ShapeEffectOverlay`'s
 * outline); this is the shared decision the other bindings map.
 *
 * @module render/picture-frame
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { getAnimationColorBaseStyle } from './animation-color-base-style';
import { hasShapeProperties, isImageLikeElement } from 'ooxml-core/pptx';

import type { CssStyleMap } from './element-style-transform';
import type { ElementAnimationState } from './animation-timeline-group';
import { buildStrokeOutline } from './stroke-outline';
import { getCompositeOuterShadowFilterCss } from './visual-effects';

/** The clip a picture box carries when it has no frame. */
export interface PictureClipStyle {
	overflow: 'hidden' | 'visible';
	borderRadius?: string | undefined;
	clipPath?: string | undefined;
}

/** How a binding styles a picture's box and, when it has a frame, its clipped surface. */
export interface PictureFrameLayers {
	/** Spread onto the picture box after the container style. */
	box: CssStyleMap;
	/**
	 * Style of the inner layer that holds the picture's pixels (the `<img>`,
	 * tile, colour wash and filter defs) and clips them, or `undefined` when
	 * the box itself clips them (no frame).
	 */
	surface: CssStyleMap | undefined;
}

/** Paint order applies to every surface, including separately animated background paint. */
export function getShapeEffectLayerPlacement(element: PptxElement): {
	beforeBody: boolean;
	afterBody: boolean;
} {
	const afterBody = isImageLikeElement(element);
	return { beforeBody: !afterBody, afterBody };
}

/** A separately animated paint layer must keep the same effect styling and visibility. */
export function getShapeEffectAnimationLayerStyle<T extends object>(
	element: PptxElement,
	visualStyle: T,
	state: ElementAnimationState,
) {
	return {
		...getAnimationColorBaseStyle(element, state),
		...visualStyle,
		...(isImageLikeElement(element) ? { clipPath: 'none', overflow: 'visible' } : {}),
		position: 'absolute',
		inset: '0',
		transformOrigin: 'center',
		visibility: state.visible === false ? 'hidden' : 'visible',
		...(state.cssAnimation ? { animation: state.cssAnimation } : {}),
		pointerEvents: 'none',
	};
}

/** The composite outer shadow a picture casts, or `undefined` without one. */
export function getPictureShadowFilter(element: PptxElement): string | undefined {
	if (!isImageLikeElement(element) || !hasShapeProperties(element)) {
		return undefined;
	}
	return getCompositeOuterShadowFilterCss(element.shapeStyle, { rotation: element.rotation });
}

/** True when the picture paints an outline or an outer shadow past its box. */
export function hasPictureFrame(element: PptxElement): boolean {
	if (!isImageLikeElement(element)) {
		return false;
	}
	return buildStrokeOutline(element) !== undefined || getPictureShadowFilter(element) !== undefined;
}

/**
 * Split a picture's clip between its box and an inner surface.
 *
 * Without a frame the box keeps `clip` exactly as before. With one, the box is
 * unclipped (`overflow: visible`) and carries the shadow `filter`, and `clip`
 * moves to a surface layer that fills the box.
 */
export function getPictureFrameLayers(
	element: PptxElement,
	clip: PictureClipStyle,
): PictureFrameLayers {
	const clipStyle: CssStyleMap = {
		overflow: clip.overflow,
		...(clip.borderRadius ? { borderRadius: clip.borderRadius } : {}),
		...(clip.clipPath ? { clipPath: clip.clipPath } : {}),
	};
	if (!hasPictureFrame(element)) {
		return { box: clipStyle, surface: undefined };
	}
	const shadow = getPictureShadowFilter(element);
	return {
		box: { overflow: 'visible', ...(shadow ? { filter: shadow } : {}) },
		surface: { position: 'absolute', inset: '0', ...clipStyle },
	};
}
