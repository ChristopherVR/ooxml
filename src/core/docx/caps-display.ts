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
