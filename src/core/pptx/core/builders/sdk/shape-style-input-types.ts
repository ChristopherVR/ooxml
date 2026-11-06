/**
 * SDK input types for shape effects, line ends and gradient rectangles,
 * kept apart from `types.ts` so that file stays a readable index of the
 * element options. Every length is in pixels, like the rest of the SDK.
 *
 * @module sdk/shape-style-input-types
 */

import type { ConnectorArrowType } from '../../types/common';

/** A rectangle as fractions (0-1) of the shape box, inset from each edge. */
export interface RelativeRectInput {
	l: number;
	t: number;
	r: number;
	b: number;
}

/** An inner shadow (`a:innerShdw`). Offsets point at the edge that is darkened. */
export interface InnerShadowInput {
	/** Default `#000000`. */
	color?: string;
	/** Blur radius in px (default 6). */
	blur?: number;
	offsetX?: number;
	offsetY?: number;
	/** 0-1 (default 0.5). */
	opacity?: number;
}

/** A glow around the shape (`a:glow`). */
export interface GlowInput {
	color: string;
	/** Glow radius in px. */
	radius: number;
	/** 0-1 (default fully opaque). */
	opacity?: number;
}

/**
 * A reflection below the shape (`a:reflection`), mirrored the way
 * PowerPoint's reflection presets are (`dir` 90 degrees, `sy` -100%,
 * aligned bottom-left).
 */
export interface ReflectionInput {
	/** Blur radius in px. */
	blur?: number;
	/** Gap between the shape and its reflection in px. */
	distance?: number;
	/** Opacity where the reflection starts, 0-1 (default 0.5). */
	startOpacity?: number;
	/** Opacity where it fades out, 0-1 (default 0). */
	endOpacity?: number;
	/** How far down the shape the fade ends, 0-1 (default 0.9). */
	endPosition?: number;
}

/** An arrowhead or other decoration on one end of an outline (`a:headEnd` / `a:tailEnd`). */
export interface LineEndInput {
	type: ConnectorArrowType;
	width?: 'sm' | 'med' | 'lg';
	length?: 'sm' | 'med' | 'lg';
}

/**
 * How a tiled picture repeats (`a:tile`). Every field is written, with
 * PowerPoint's defaults for the omitted ones.
 */
export interface PictureTileInput {
	/** Offset of the tiles from where they start, in px (`@tx`, default 0). */
	offsetX?: number;
	/** Offset of the tiles from where they start, in px (`@ty`, default 0). */
	offsetY?: number;
	/** Tile width as a fraction of the picture's own width, above 0 (`@sx`, default 1). */
	scaleX?: number;
	/** Tile height as a fraction of the picture's own height, above 0 (`@sy`, default 1). */
	scaleY?: number;
	/** Mirror every other tile (`@flip`, default `'none'`). */
	flip?: 'none' | 'x' | 'y' | 'xy';
	/** The corner, edge or centre the tiles start from (`@algn`, default `'tl'`). */
	alignment?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
}
