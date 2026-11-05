/**
 * Whether a CSS `font` shorthand will paint with the face it names, rather
 * than with a stand-in, so a measurement of it can be trusted and cached.
 *
 * `document.fonts.check` alone cannot say: Chromium answers `true` for a
 * family that matches no `@font-face` at all (the "every matched face is
 * loaded" condition holds vacuously), which is also what a Google web font
 * looks like before its stylesheet has arrived. So a family with no face is
 * settled only when the machine itself provides it (the canvas width test of
 * {@link isFontFamilyInstalledLocally}), and a family with faces is settled
 * only when the faces its text needs have loaded. Asking also starts the
 * load, which a canvas `font` assignment does for a face the page has not
 * used yet but a page that never drew with it has not requested.
 *
 * @module render/font-settled
 */
import { isFontFamilyInstalledLocally } from './google-webfonts';

/** Generic families are always provided by the browser. */
const GENERIC_FAMILIES = new Set([
	'serif',
	'sans-serif',
	'monospace',
	'cursive',
	'fantasy',
	'system-ui',
	'ui-serif',
	'ui-sans-serif',
	'ui-monospace',
	'ui-rounded',
	'emoji',
	'math',
	'fangsong',
]);

const installedCache = new Map<string, boolean>();

/** The first family of a `font` shorthand (`italic 700 40px "A B", C`), unquoted. */
export function firstFontFamily(font: string): string | undefined {
	const list = /\d(?:px|pt|em|rem|%)(?:\s*\/\s*[^\s]+)?\s+(.+)$/u.exec(font)?.[1];
	const first = list
		?.split(',')[0]
		?.trim()
		.replace(/^["']|["']$/gu, '');
	return first ? first : undefined;
}

/** Whether `font`'s first family paints with its own face, as described in the module comment. */
export function isFontSettled(font: string): boolean {
	return fontState(font) === 'settled';
}

/**
 * `settled`: the font paints with its own face; `loading`: a web font face of
 * the family exists but has not loaded (its own box is about to be known);
 * `missing`: no face and not installed, so the browser paints a fallback (a
 * face may still arrive).
 */
export function fontState(font: string): 'settled' | 'loading' | 'missing' {
	if (typeof document === 'undefined') {
		return 'settled';
	}
	const family = firstFontFamily(font);
	if (!family || GENERIC_FAMILIES.has(family.toLowerCase())) {
		return 'settled';
	}
	const fonts = document.fonts;
	let hasFace = false;
	fonts?.forEach((face) => {
		if (face.family.replace(/^["']|["']$/gu, '') === family) {
			hasFace = true;
		}
	});
	if (hasFace) {
		if (fonts.check(font, 'x')) {
			return 'settled';
		}
		void fonts.load(font, 'x').catch(() => undefined);
		return 'loading';
	}
	let installed = installedCache.get(family);
	if (installed === undefined) {
		installed = isFontFamilyInstalledLocally(family);
		installedCache.set(family, installed);
	}
	return installed ? 'settled' : 'missing';
}

/** Forget which families are installed, for a face that has since been added. */
export function resetInstalledFontCache(): void {
	installedCache.clear();
}
