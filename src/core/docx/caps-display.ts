/**
 * The text Word draws for All caps (`w:caps`) and Small caps (`w:smallCaps`). Word 16 drops the
 * accents of the letters the formatting capitalises when the run's language is French (France):
 * "Procédure" in All caps reads PROCEDURE in fr-FR but PROCÉDURE in fr-CA and en-US, and capitals
 * typed as such keep theirs. Print Layout and the PDF export share this projection. DOM-free.
 */

const COMBINING_MARKS = /\p{M}+/gu;
const LOWERCASE = /\p{Ll}/u;

/** Whether capitals made by caps formatting lose their accents in `language` (only fr-FR). */
export function capsDropAccents(language: string | undefined): boolean {
	return language?.toLowerCase() === 'fr-fr';
}

const withoutAccents = (text: string): string =>
	text.normalize('NFD').replace(COMBINING_MARKS, '').normalize('NFC');

/** `text` in All caps for a run in `language`. */
export function allCapsText(text: string, language: string | undefined): string {
	if (!capsDropAccents(language)) return text.toUpperCase();
	let out = '';
	// Whether the last letter was capitalised by the formatting: its combining accents go too.
	let made = false;
	for (const char of text) {
		if (/^\p{M}$/u.test(char)) {
			if (!made) out += char;
			continue;
		}
		made = LOWERCASE.test(char);
		out += made ? withoutAccents(char.toUpperCase()) : char.toUpperCase();
	}
	return out;
}

/**
 * `text` for a Small caps run in `language`: unchanged, except that in fr-FR an accented lowercase
 * letter (which small caps draws as a capital) loses its accent. Only precomposed letters change,
 * one character for one, so source offsets are kept.
 */
export function smallCapsText(text: string, language: string | undefined): string {
	if (!capsDropAccents(language)) return text;
	let out = '';
	for (const char of text) {
		const base = LOWERCASE.test(char) ? withoutAccents(char) : char;
		out += base.length === char.length ? base : char;
	}
	return out;
}
