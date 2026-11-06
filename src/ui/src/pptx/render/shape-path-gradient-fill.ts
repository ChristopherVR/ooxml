/**
 * `path="shape"` gradient paint for a preset, resolved through the shared
 * scaled-outline field (see `ooxml-core` `shape-path-gradient`): the outline
 * the shape is drawn with, scaled about the gradient focus, as an SVG image a
 * `background-image` can take.
 *
 * @module shape-path-gradient-fill
 */
import { buildOutlinePathGradientSvg, closedSubpaths } from 'ooxml-core/diagram';
import { evaluatePresetShape, getShapeType } from 'ooxml-core/pptx';

import type { RectPathGradientStop } from './path-gradient-rect';

/** The geometry a shape-path gradient is painted on. */
export interface ShapePathGradientGeometry {
	shapeType: string;
	width: number;
	height: number;
	adjustments?: Record<string, number> | undefined;
}

const CACHE_LIMIT = 64;
const cache = new Map<string, string | null>();

function toCssUrl(svg: string): string {
	return `url("data:image/svg+xml,${encodeURIComponent(svg).replace(/[()]/gu, (c) => (c === '(' ? '%28' : '%29'))}")`;
}

/** The outline path of the shape body: the first sub-path that paints the plain fill. */
function bodyPath(geometry: ShapePathGradientGeometry): string | undefined {
	const { shapeType, width, height, adjustments } = geometry;
	const evaluated =
		evaluatePresetShape(shapeType, width, height, adjustments) ??
		evaluatePresetShape(getShapeType(shapeType), width, height, adjustments);
	const body = evaluated?.paths.find((path) => !path.fill || path.fill === 'norm');
	return body ? closedSubpaths(body.d) : undefined;
}

/**
 * The CSS `background-image` for a `path="shape"` gradient on a preset,
 * or `undefined` when the outline cannot be scaled (holes, an unknown preset),
 * in which case the caller keeps its approximation.
 *
 * @param focus - the gradient focus as fractions (0-1) of the box
 */
export function buildOutlineShapePathGradient(
	stops: readonly RectPathGradientStop[],
	geometry: ShapePathGradientGeometry,
	focus: { x: number; y: number },
): string | undefined {
	if (!(geometry.width > 0) || !(geometry.height > 0)) return undefined;
	const key = JSON.stringify([stops, geometry, focus]);
	const known = cache.get(key);
	if (known !== undefined) return known ?? undefined;
	const d = bodyPath(geometry);
	const svg = d
		? buildOutlinePathGradientSvg(stops, d, geometry.width, geometry.height, focus.x, focus.y)
		: undefined;
	const result = svg ? toCssUrl(svg) : null;
	if (cache.size >= CACHE_LIMIT) cache.delete(cache.keys().next().value as string);
	cache.set(key, result);
	return result ?? undefined;
}
