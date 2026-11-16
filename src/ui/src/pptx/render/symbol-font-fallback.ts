/**
 * Bullet markers and text in symbol fonts on devices that lack the font.
 *
 * PowerPoint's square, star, arrow and check bullets are Wingdings characters
 * (`§`, `v`, `Ø`, `ü`). Where Wingdings is installed the browser paints the
 * glyph; on iPadOS, Android or Linux the character falls back to an ordinary
 * font and the bullet shows as a literal `§` (issue #290). When the browser
 * cannot paint the glyph in its font, the marker is replaced by its Unicode
 * equivalent and painted in the paragraph's typeface instead. Runs in a
 * symbol font (a Wingdings checkmark column, Symbol Greek, an `a:sym` glyph)
 * get the same treatment.
 */

import { isSymbolEncodedFont, symbolCharToUnicode, symbolTextToUnicode } from 'ooxml-core/text';

import type { ParagraphBulletResult } from './bullet-list';
import type { ParagraphRun } from './paragraph-types';
import type { RunStyle } from './text-run-style';

/** Reports whether `family` really paints `char`, rather than a fallback font. */
export type SymbolGlyphProbe = (family: string, char: string) => boolean;

interface WidthContext {
	font: string;
	measureText(text: string): { width: number };
}

let probeContext: WidthContext | null | undefined;
const probeCache = new Map<string, boolean>();

function widthContext(): WidthContext | null {
	if (probeContext === undefined) {
		try {
			probeContext =
				typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
		} catch {
			probeContext = null;
		}
	}
	return probeContext;
}

/**
 * Canvas probe: a glyph the font lacks is drawn by the fallback, so its advance
 * matches the fallback's own in two unrelated generic families. Without a
 * canvas (server rendering, tests) the font counts as missing, as it is on most
 * devices.
 */
export const canvasSymbolGlyphProbe: SymbolGlyphProbe = (family, char) => {
	const key = `${family.toLowerCase()}\u0000${char}`;
	const cached = probeCache.get(key);
	if (cached !== undefined) return cached;
	const ctx = widthContext();
	let painted = false;
	if (ctx) {
		const quoted = `"${family.replace(/["\\]/g, '')}"`;
		painted = ['monospace', 'serif'].some((generic) => {
			ctx.font = `72px ${generic}`;
			const fallback = ctx.measureText(char).width;
			ctx.font = `72px ${quoted}, ${generic}`;
			return Math.abs(ctx.measureText(char).width - fallback) > 0.01;
		});
	}
	probeCache.set(key, painted);
	return painted;
};

/**
 * The bullet as this device can show it: unchanged unless its font is a symbol
 * font the browser cannot paint the marker in, in which case the marker is its
 * Unicode equivalent (a plain bullet when none is known) and the bullet drops
 * its font so the marker is drawn in the paragraph's typeface.
 */
export function displayableSymbolBullet(
	bullet: ParagraphBulletResult | undefined,
	probe: SymbolGlyphProbe = canvasSymbolGlyphProbe,
): ParagraphBulletResult | undefined {
	const family = bullet?.fontFamily;
	if (!bullet || bullet.isNumbered || bullet.picture || !family || !isSymbolEncodedFont(family)) {
		return bullet;
	}
	const marker = bullet.marker;
	const glyphs = [...marker].filter((char) => char.trim() !== '');
	if (glyphs.length === 0 || glyphs.every((char) => probe(family, char))) {
		return bullet;
	}
	const { fontFamily: _symbolFont, ...rest } = bullet;
	return { ...rest, marker: symbolTextToUnicode(family, marker) ?? '\u2022' };
}

function primaryFamily(fontFamily: unknown): string | undefined {
	if (typeof fontFamily !== 'string') return undefined;
	const first = fontFamily
		.split(',')[0]
		?.trim()
		.replace(/^["']|["']$/g, '');
	return first || undefined;
}

/**
 * A piece of text as this device can show it. Each glyph the symbol font cannot
 * paint becomes its Unicode equivalent (one UTF-16 unit, so offsets into the
 * run stay valid); a glyph with no known equivalent is left as it is.
 */
function displayablePiece<T extends { text: string; style?: RunStyle }>(
	piece: T,
	fallbackFamily: string,
	probe: SymbolGlyphProbe,
): T {
	const family = primaryFamily(piece.style?.fontFamily);
	if (!family || !isSymbolEncodedFont(family) || !piece.text) return piece;
	let changed = false;
	let text = '';
	for (const char of piece.text) {
		const mapped =
			char.trim() === '' || probe(family, char) ? undefined : symbolCharToUnicode(family, char);
		changed ||= mapped !== undefined;
		text += mapped ?? char;
	}
	if (!changed) return piece;
	return { ...piece, text, style: { ...piece.style, fontFamily: fallbackFamily } };
}

/**
 * The paragraph's runs as this device can show them: text in a symbol font the
 * browser cannot paint is rewritten in Unicode and drawn in `fallbackFamily`
 * (the text body's typeface). Runs in ordinary fonts are returned as they are.
 */
export function displayableSymbolRuns(
	runs: ParagraphRun[],
	fallbackFamily: string,
	probe: SymbolGlyphProbe = canvasSymbolGlyphProbe,
): ParagraphRun[] {
	return runs.map((run) => {
		let next = displayablePiece(run, fallbackFamily, probe);
		for (const key of ['scriptRuns', 'underlineWordPieces'] as const) {
			const pieces = run[key];
			if (!pieces) continue;
			const shown = pieces.map((piece) => displayablePiece(piece, fallbackFamily, probe));
			if (shown.some((piece, index) => piece !== pieces[index])) {
				next = { ...next, [key]: shown };
			}
		}
		return next;
	});
}
