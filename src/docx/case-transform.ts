export type CaseMode = 'sentence' | 'lower' | 'upper' | 'title' | 'toggle';

/** Characters that make up a word, shared with the editor's caret-word lookup. */
export const WORD = /[\p{L}\p{N}'’]/u;

const swap = (char: string): string => {
	const upper = char.toLocaleUpperCase();
	return char === upper ? char.toLocaleLowerCase() : upper;
};

/**
 * Applies a Word "Change Case" mode to `text`. `atSentenceStart` says whether the text begins a
 * sentence; it matters only for sentence case, which capitalises the first letter after `. ! ?`.
 */
export function transformCase(text: string, mode: CaseMode, atSentenceStart = true): string {
	if (mode === 'lower') return text.toLocaleLowerCase();
	if (mode === 'upper') return text.toLocaleUpperCase();
	if (mode === 'toggle') return [...text].map(swap).join('');
	let start = mode === 'sentence' ? atSentenceStart : true;
	let previousWord = false;
	let out = '';
	for (const char of text) {
		const isWord = WORD.test(char);
		if (mode === 'title') {
			out += isWord && !previousWord ? char.toLocaleUpperCase() : char;
		} else if (isWord && start) {
			out += char.toLocaleUpperCase();
			start = false;
		} else {
			out += isWord ? char.toLocaleLowerCase() : char;
			if (/[.!?]/.test(char)) start = true;
			else if (isWord) start = false;
		}
		previousWord = isWord;
	}
	return out;
}
