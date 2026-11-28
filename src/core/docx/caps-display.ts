/**
 * The text Word draws for All caps (`w:caps`) and Small caps (`w:smallCaps`). Word 16 drops the
 * accents of the letters the formatting capitalises when the run's language is French (France):
 * "Procédure" in All caps reads PROCEDURE in fr-FR but PROCÉDURE in fr-CA and en-US, and capitals
 * typed as such keep theirs. Print Layout and the PDF export share this projection. DOM-free.
 */

const withoutAccents = (text: string): string =>
	text
		.normalize('NFD')
		.replace(/\p{M}+/gu, '')
		.normalize('NFC');

/** Whether capitals made by caps formatting lose their accents in `language` (only fr-FR). */
const dropsAccents = (language: string | undefined) => language?.toLowerCase() === 'fr-fr';

/** `text` in All caps for a run in `language`. */
export function allCapsText(text: string, language: string | undefined): string {
	// A lowercase letter, with any combining accents after it, loses them as it is capitalised.
	return (
		dropsAccents(language) ? text.replace(/\p{Ll}\p{M}*/gu, withoutAccents) : text
	).toUpperCase();
}

/**
 * `text` for a Small caps run in `language`: unchanged, except that in fr-FR an accented lowercase
 * letter (which small caps draws as a capital) loses its accent. Only precomposed letters change,
 * one character for one, so source offsets are kept.
 */
export function smallCapsText(text: string, language: string | undefined): string {
	return dropsAccents(language)
		? text.replace(/\p{Ll}/gu, (char) => {
				const base = withoutAccents(char);
				return base.length === char.length ? base : char;
			})
		: text;
}

/** Word draws small caps at 0.8 of the run's size (measured, Word 16). */
export const SMALL_CAPS_SCALE = 0.8;

/**
 * Splits `text` for drawing in small caps: lowercase letters and the spaces among them are
 * capitalised and drawn at {@link SMALL_CAPS_SCALE}; capitals and other characters keep the full
 * size. Word draws a space like a lowercase letter, at the reduced size, even between capitals
 * (a space is 2.23 pt in 11 pt Times New Roman bold small caps, 2.75 pt without).
 */
export function smallCapsParts(text: string): { text: string; small: boolean }[] {
	return [...text.matchAll(/(?<small>(?:[\p{Ll} ]\p{M}*)+)|(?<full>[^\p{Ll} ]+)/gu)].map(
		(match) => {
			const small = match.groups?.small !== undefined;
			return { text: match[0].toUpperCase(), small };
		},
	);
}
