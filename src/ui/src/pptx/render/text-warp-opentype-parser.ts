import type { Font } from 'opentype.js';

export type OpenTypeFontParser = (buffer: ArrayBuffer) => Font;

type OpenTypeModuleShape = {
	parse?: OpenTypeFontParser;
	default?: { parse?: OpenTypeFontParser };
};

let loadedModule: OpenTypeModuleShape | undefined;
let loading: Promise<void> | undefined;

/** True once {@link parseOpenTypeFont} can parse without an explicit module. */
export function isOpenTypeFontParserReady(): boolean {
	return loadedModule !== undefined;
}

/**
 * Load `opentype.js` on first demand and share that one import. It is only
 * needed to warp WordArt envelope glyphs along a font file's real outlines, so
 * it stays out of the initial bundle. A failed import is forgotten, so a later
 * demand retries it.
 */
export function preloadOpenTypeFontParser(): Promise<void> {
	loading ??= import('opentype.js')
		.then((module) => {
			const resolved = module as unknown as OpenTypeModuleShape;
			if (!resolved.parse && !resolved.default?.parse) {
				throw new Error('opentype.js does not expose a font parser');
			}
			loadedModule = resolved;
		})
		.catch((error: unknown) => {
			loading = undefined;
			throw error;
		});
	return loading;
}

/**
 * Parse a font file with opentype.js as either its ESM namespace or its
 * CommonJS default object. Without an explicit `module`, the parser loaded by
 * {@link preloadOpenTypeFontParser} is used, and parsing before it is ready throws.
 */
export function parseOpenTypeFont(
	buffer: ArrayBuffer,
	module: OpenTypeModuleShape | undefined = loadedModule,
): Font {
	if (!module) {
		throw new Error('opentype.js is not loaded yet; await preloadOpenTypeFontParser() first');
	}
	const parser = module.parse ?? module.default?.parse;
	if (!parser) {
		throw new Error('opentype.js does not expose a font parser');
	}
	return parser(buffer);
}
