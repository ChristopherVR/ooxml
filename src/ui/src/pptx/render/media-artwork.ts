import { normalizeStShapeType } from 'ooxml-core/geometry';
import type { PptxElement } from 'ooxml-core/pptx';

import { getResolvedShapeClipPath, isIdentityRectClip } from './shape-geometry';

/** Clip the artwork itself so the media transport can still appear beneath it. */
export function mediaArtworkClipPath(element: PptxElement, fullscreen = false): string | undefined {
	if (element.type !== 'media' || fullscreen || isIdentityRectClip(element)) return undefined;
	return normalizeStShapeType(element.shapeType) === 'ellipse'
		? 'ellipse(50% 50%)'
		: getResolvedShapeClipPath(element);
}
