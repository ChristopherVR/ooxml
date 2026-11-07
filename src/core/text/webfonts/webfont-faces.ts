/**
 * The faces a document draws in webfont families, from format-neutral text runs. Every product
 * reduces its own model to {@link WebfontRun}s (PowerPoint's shapes and tables, Word's runs,
 * Excel's cell fonts, Visio's characters), and this decides which faces to load and which
 * characters each draws, so the right `unicode-range` subset is requested.
 *
 * Moved from PowerPoint's `deck-webfont-readiness.ts` so every product shares it.
 *
 * @module text/webfonts/webfont-faces
 */
import { resolveFontItalic, resolveFontWeight } from './font-weight-names';
import { servedGoogleFontsFamily } from './google-fonts-request';

/** One run of text: its family, emphasis and characters. */
export interface WebfontRun {
	family: string | undefined;
	bold?: boolean | undefined;
	italic?: boolean | undefined;
	text: string;
}

/** One face to load: CSS `font` shorthand, its served family and the characters drawn in it. */
export interface WebfontFace {
	family: string;
	font: string;
	chars: Set<string>;
}

/** Characters collected per face for its `unicode-range` subset (enough for any script mix). */
const MAX_SAMPLE_CHARS = 256;

const quoteFamily = (family: string) => `"${family.replace(/[\\"]/gu, '\\$&')}"`;

/** Every distinct family the runs name. */
export function collectRunFamilies(runs: Iterable<WebfontRun>): Set<string> {
	const families = new Set<string>();
	for (const run of runs) {
		const family = run.family?.trim();
		if (family) families.add(family);
	}
	return families;
}

/**
 * The faces the runs draw in `candidates`, grouped by served family, and the candidates the
 * catalogue does not serve.
 */
export function collectWebfontFaces(
	runs: Iterable<WebfontRun>,
	candidates: readonly string[],
): { faces: Map<string, WebfontFace[]>; unavailable: string[] } {
	const wanted = new Set(candidates);
	const unavailable = candidates.filter((family) => servedGoogleFontsFamily(family) === null);
	const faces = new Map<string, WebfontFace[]>();
	const byFont = new Map<string, WebfontFace>();
	for (const run of runs) {
		const family = run.family?.trim();
		if (!family || !wanted.has(family)) continue;
		const served = servedGoogleFontsFamily(family);
		if (served === null) continue;
		const weight = resolveFontWeight(family, run.bold);
		const style = resolveFontItalic(family, run.italic) ? 'italic ' : '';
		const font = `${style}${weight} 16px ${quoteFamily(served)}`;
		let face = byFont.get(font);
		if (!face) {
			face = { family: served, font, chars: new Set() };
			byFont.set(font, face);
			const list = faces.get(served);
			if (list) list.push(face);
			else faces.set(served, [face]);
		}
		for (const char of run.text) {
			if (face.chars.size >= MAX_SAMPLE_CHARS) break;
			if (char.trim() !== '') face.chars.add(char);
		}
	}
	return { faces, unavailable };
}
