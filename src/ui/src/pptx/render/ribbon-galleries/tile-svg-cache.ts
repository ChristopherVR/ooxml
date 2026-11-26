/**
 * Tile previews are pure functions of their inputs, and a ribbon rebuilds every gallery's
 * descriptor on each state update: the Shape Styles and Shape Effects galleries alone paint about
 * 180 tiles per rebuild, dozens of times while a deck opens. The markup is cached by its inputs,
 * up to a bound, since a theme change brings in a new set of colours.
 *
 * @module render/ribbon-galleries/tile-svg-cache
 */

const LIMIT = 2048;
const cache = new Map<string, string>();

/** The markup `paint` produces for `key`, painted once per key. */
export function cachedTileSvg(key: string, paint: () => string): string {
	const known = cache.get(key);
	if (known !== undefined) {
		return known;
	}
	const svg = paint();
	if (cache.size >= LIMIT) {
		cache.clear();
	}
	cache.set(key, svg);
	return svg;
}
