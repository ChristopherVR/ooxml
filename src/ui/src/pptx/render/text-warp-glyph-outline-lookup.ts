/**
 * Glyph-outline lookup over a {@link GlyphOutlineFontCache}: extracts one
 * glyph's outline commands from a parsed font file, and builds the
 * `getGlyphOutline` callback `buildGlyphEnvelopeBlock` accepts, falling back
 * to tracing the browser's own rendering when no font file is available.
 *
 * Kept apart from `text-warp-outline-font-cache.ts` (which owns parsing,
 * caching and the on-demand parser load) and re-exported from it, so callers
 * keep importing these from the cache module.
 */
import type { Font } from 'opentype.js';

import { DEFAULT_TEXT_FONT_SIZE } from '../constants';
import type { GlyphOutlineCommand } from './text-warp-glyph-outline';
import { traceGlyphOutlineCommands } from './text-warp-glyph-trace';
import type { GlyphOutlineFontCache } from './text-warp-outline-font-cache';

/** Characters with no visible glyph outline (a legitimate "nothing to draw" result). */
const WHITESPACE_RE = /^\s$/u;

/**
 * Extract one glyph's outline commands from `font`, positioned at `(x, y)`
 * (the glyph's own baseline origin, matching `EnvelopeGlyphPlacement.x`/`.y`)
 * and scaled to `fontSizePx`.
 *
 * Returns `[]` for whitespace (no visible outline; still a resolvable
 * result, distinct from "unobtainable") and `undefined` when `font` cannot
 * shape `char` at all, so the caller falls back to the affine transform.
 */
export function extractGlyphOutlineCommands(
	font: Font,
	char: string,
	x: number,
	y: number,
	fontSizePx: number,
): GlyphOutlineCommand[] | undefined {
	if (!char) {
		return undefined;
	}
	if (WHITESPACE_RE.test(char)) {
		return [];
	}
	try {
		if (!font.hasChar(char)) {
			return undefined;
		}
		const path = font.getPath(char, x, y, fontSizePx);
		return path.commands.map((cmd): GlyphOutlineCommand => {
			switch (cmd.type) {
				case 'M':
				case 'L':
					return { type: cmd.type, x: cmd.x, y: cmd.y };
				case 'Q':
					return { type: 'Q', x1: cmd.x1, y1: cmd.y1, x: cmd.x, y: cmd.y };
				case 'C':
					return { type: 'C', x1: cmd.x1, y1: cmd.y1, x2: cmd.x2, y2: cmd.y2, x: cmd.x, y: cmd.y };
				case 'Z':
				default:
					return { type: 'Z' };
			}
		});
	} catch {
		return undefined;
	}
}

/** A minimal font-spec shape, structurally compatible with `EnvelopeFontSpec`. */
export interface GlyphOutlineCacheFontSpec {
	fontFamily?: string;
	fontSizePx?: number;
	bold?: boolean;
	italic?: boolean;
}

/**
 * Build the `getGlyphOutline` callback `buildGlyphEnvelopeBlock` (in
 * `text-warp-envelope-layout.ts`) accepts: looks up `cache` for the segment's
 * resolved font and, if found, extracts that glyph's outline from the font
 * file; otherwise traces the browser's own rendering of the glyph (see
 * `text-warp-glyph-trace.ts`). `undefined` only without a DOM canvas.
 */
export function createGlyphOutlineLookup(
	cache: GlyphOutlineFontCache,
): (
	char: string,
	font: GlyphOutlineCacheFontSpec,
	x: number,
	y: number,
) => GlyphOutlineCommand[] | undefined {
	return (char, font, x, y) => {
		// Whitespace has no outline in any font, so it never asks for one.
		if (WHITESPACE_RE.test(char)) {
			return [];
		}
		const parsed = cache.get(font.fontFamily, font.bold, font.italic);
		if (parsed) {
			const size =
				font.fontSizePx && font.fontSizePx > 0 ? font.fontSizePx : DEFAULT_TEXT_FONT_SIZE;
			const commands = extractGlyphOutlineCommands(parsed, char, x, y, size);
			if (commands) {
				return commands;
			}
		}
		// No font file (or it lacks this glyph): trace the glyph the browser
		// itself renders for this font, so it still gets the exact warp.
		return traceGlyphOutlineCommands(char, font, x, y);
	};
}
