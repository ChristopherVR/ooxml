/**
 * The fill of a silhouette shape (see `silhouette-effects`) whose fill is a
 * gradient, pattern or picture: the container's filter paints its shadow and
 * glow from the painted alpha, so the fill must not be clipped on the container.
 * It is a child layer carrying the same CSS background and the clip-path of the
 * geometry; a child's clip applies before the container's filter.
 *
 * @module render/silhouette-fill-layer
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import { resolveComputedFill } from './fill-style';
import { resolveBaseShapeGeometry } from './shape-geometry-base';
import { silhouetteFillsLayer } from './silhouette-effects';

/** The camelCase style of the fill layer, to be rendered as an inert absolute child. */
export interface SilhouetteFillLayer {
	readonly style: Readonly<Record<string, string>>;
	/** The same declarations as a `style` attribute value. */
	readonly cssText: string;
}

/** The fill layer of a shape, or `undefined` when its fill is not painted that way. */
export function getSilhouetteFillLayer(element: PptxElement): SilhouetteFillLayer | undefined {
	if (!silhouetteFillsLayer(element) || !hasShapeProperties(element) || !element.shapeStyle) {
		return undefined;
	}
	const geometry = resolveBaseShapeGeometry(element);
	if (geometry.kind !== 'clipPath') {
		return undefined;
	}
	const fill = resolveComputedFill(element.shapeStyle, element);
	const style: Record<string, string> = {
		position: 'absolute',
		inset: '0',
		pointerEvents: 'none',
		clipPath: geometry.clipPath,
	};
	for (const key of [
		'backgroundColor',
		'backgroundImage',
		'backgroundRepeat',
		'backgroundSize',
		'backgroundPosition',
	] as const) {
		const value = fill[key];
		if (typeof value === 'string') {
			style[key] = value;
		}
	}
	const cssText = Object.entries(style)
		.map(([k, v]) => `${k.replace(/[A-Z]/gu, (c) => `-${c.toLowerCase()}`)}:${v}`)
		.join(';');
	return { style, cssText };
}
