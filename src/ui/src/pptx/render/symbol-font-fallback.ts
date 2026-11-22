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
import { withSymbolFallback } from '../../fonts/symbol-fallback';

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
	// A font can arrive after the first render (embedded or host supplied).
	// Never retain a negative result across that load.
	if (painted) probeCache.set(key, true);
	return painted;
};

/**
 * The bullet as this device can show it: unchanged unless its font is a symbol
 * font the browser cannot paint the marker in, in which case the marker is its
 * Unicode equivalent (a plain bullet when none is known), drawn in the paragraph
 * typeface with the bundled symbol font as its last resort.
 */
export function displayableSymbolBullet(
	bullet: ParagraphBulletResult | undefined,
	probe: SymbolGlyphProbe = canvasSymbolGlyphProbe,
	paragraphFamily?: string,
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
	// A marker outside the legacy byte and private-use ranges already has a
	// Unicode meaning. Preserve it and let the browser fall through the font stack.
	if (glyphs.every((char) => !/[\u0000-\u00ff\p{Co}]/u.test(char))) {
		return { ...bullet, fontFamily: withSymbolFallback(family) };
	}
	return {
		...bullet,
		marker: symbolTextToUnicode(family, marker) ?? '\u2022',
		fontFamily: withSymbolFallback(paragraphFamily ?? 'sans-serif'),
	};
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
 * A piece of text as this device can show it. `inherited` is the family the
 * piece paints in when it names none of its own (its run's, or the text
 * body's). Each glyph the symbol font cannot paint becomes its Unicode
 * equivalent (one UTF-16 unit, so offsets into the run stay valid); a glyph
 * with no known equivalent is left as it is.
 */
function displayablePiece<T extends { text: string; style?: RunStyle }>(
	piece: T,
	inherited: string | undefined,
	fallbackFamily: string,
	probe: SymbolGlyphProbe,
): T {
	const family = primaryFamily(piece.style?.fontFamily ?? inherited);
	if (!family || !isSymbolEncodedFont(family) || !piece.text) return piece;
	let changed = false;
	let text = '';
	for (const char of piece.text) {
		const mapped =
			char.trim() === '' || probe(family, char) ? undefined : symbolCharToUnicode(family, char);
		// Selection and measurement offsets refer to the source UTF-16 run.
		const replacement = mapped?.length === char.length ? mapped : undefined;
		changed ||= replacement !== undefined;
		text += replacement ?? char;
	}
	if (!changed) return piece;
	return {
		...piece,
		text,
		style: { ...piece.style, fontFamily: withSymbolFallback(fallbackFamily) },
	};
}

function displayablePieces<T extends { text: string; style?: RunStyle }>(
	pieces: T[] | undefined,
	inherited: string | undefined,
	fallbackFamily: string,
	probe: SymbolGlyphProbe,
): T[] | undefined {
	if (!pieces) return undefined;
	const shown = pieces.map((piece) => displayablePiece(piece, inherited, fallbackFamily, probe));
	return shown.some((piece, index) => piece !== pieces[index]) ? shown : undefined;
}

/**
 * The paragraph's runs as this device can show them: text in a symbol font the
 * browser cannot paint is rewritten in Unicode and drawn in `fallbackFamily`.
 * `bodyFamily` is the text body's own family, which runs that name no font of
 * their own inherit (a text box set wholly in Wingdings). Runs in ordinary
 * fonts are returned as they are.
 */
export function displayableSymbolRuns(
	runs: ParagraphRun[],
	fallbackFamily: string,
	probe: SymbolGlyphProbe = canvasSymbolGlyphProbe,
	bodyFamily?: string,
): ParagraphRun[] {
	return runs.map((run) => {
		const runFamily = typeof run.style.fontFamily === 'string' ? run.style.fontFamily : bodyFamily;
		let next = displayablePiece(run, bodyFamily, fallbackFamily, probe);
		for (const key of ['scriptRuns', 'underlineWordPieces'] as const) {
			const shown = displayablePieces(run[key], runFamily, fallbackFamily, probe);
			if (shown) next = { ...next, [key]: shown };
		}
		if (run.tabLines) {
			let changed = false;
			const tabLines = run.tabLines.map((line) => {
				const pieces = displayablePieces(line.pieces, runFamily, fallbackFamily, probe);
				changed ||= pieces !== undefined;
				return pieces ? { ...line, pieces } : line;
			});
			if (changed) next = { ...next, tabLines };
		}
		return next;
	});
}
