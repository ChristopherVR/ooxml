import type { ViewerProofingOptions } from './viewer-options';

/**
 * AutoCorrect-as-you-type engine behind Options > Proofing, mirroring
 * PowerPoint's replacements: initial-capitals fixes, sentence and day-name
 * capitalization, smart quotes, double-hyphen to dash, fraction glyphs, and
 * superscript-style ordinals. Bindings run `applyAutoCorrect` over committed
 * text (word boundaries, blur, or Enter), never on every keystroke.
 */

const DAY_NAMES = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];

const FRACTIONS: ReadonlyArray<readonly [string, string]> = [
	['1/2', '½'],
	['1/4', '¼'],
	['3/4', '¾'],
];

const ORDINAL_SUFFIXES: Record<string, string> = {
	st: 'ˢᵗ',
	nd: 'ⁿᵈ',
	rd: 'ʳᵈ',
	th: 'ᵗʰ',
};

function fixTwoInitialCapitals(word: string): string {
	if (/^[A-Z]{2}[a-z]/.test(word) && !/^[A-Z]{3,}$/.test(word)) {
		const second = word.charAt(1).toLowerCase();
		return `${word.charAt(0)}${second}${word.slice(2)}`;
	}
	return word;
}

function capitalizeDayName(word: string): string {
	const bare = word.toLowerCase();
	if (DAY_NAMES.includes(bare) && word === bare) {
		return `${bare.charAt(0).toUpperCase()}${bare.slice(1)}`;
	}
	return word;
}

function replaceFractions(word: string): string {
	for (const [plain, glyph] of FRACTIONS) {
		if (word === plain) {
			return glyph;
		}
	}
	return word;
}

function replaceOrdinal(word: string): string {
	const match = /^(\d+)(st|nd|rd|th)$/i.exec(word);
	if (!match || match[1] === undefined || match[2] === undefined) {
		return word;
	}
	const superscript = ORDINAL_SUFFIXES[match[2].toLowerCase()];
	return superscript ? `${match[1]}${superscript}` : word;
}

function transformWord(word: string, proofing: ViewerProofingOptions): string {
	let result = word;
	if (proofing.autoCorrectTwoInitialCapitals) {
		result = fixTwoInitialCapitals(result);
	}
	if (proofing.autoCorrectCapitalizeDayNames) {
		result = capitalizeDayName(result);
	}
	if (proofing.autoCorrectFractions) {
		result = replaceFractions(result);
	}
	if (proofing.autoCorrectOrdinals) {
		result = replaceOrdinal(result);
	}
	return result;
}

function applySmartQuotes(text: string, lead = ''): string {
	let result = '';
	for (let i = 0; i < text.length; i += 1) {
		const char = text.charAt(i);
		if (char !== '"' && char !== "'") {
			result += char;
			continue;
		}
		const previous = i === 0 ? lead.slice(-1) : text.charAt(i - 1);
		const opening = previous === '' || /[\s([{‘“-]/.test(previous);
		if (char === '"') {
			result += opening ? '“' : '”';
		} else {
			result += opening ? '‘' : '’';
		}
	}
	return result;
}

function applyHyphensToDash(text: string): string {
	return text.replace(/(\S)--(?=\S)/g, '$1—').replace(/ -- /g, ' – ');
}

/** Capitalize sentence starts in `text`; `lead` is the text before it, which is not changed. */
function capitalizeSentences(text: string, lead = ''): string {
	return (lead + text)
		.replace(/(^|[.!?]\s+)([a-z])/g, (match, head: string, letter: string, offset: number) =>
			offset + match.length - 1 >= lead.length ? `${head}${letter.toUpperCase()}` : match,
		)
		.slice(lead.length);
}
/**
 * Apply every enabled AutoCorrect rule to a run of text. `lead` is the text that
 * precedes the run (default: none, so the run starts a sentence); it only decides
 * quote direction and whether the run's first letter starts a sentence.
 */
export function applyAutoCorrect(text: string, proofing: ViewerProofingOptions, lead = ''): string {
	if (text.length === 0) {
		return text;
	}
	let result = text;
	if (proofing.autoCorrectSmartQuotes) {
		result = applySmartQuotes(result, lead);
	}
	if (proofing.autoCorrectHyphensToDash) {
		result = applyHyphensToDash(result);
	}
	const needsWordPass =
		proofing.autoCorrectTwoInitialCapitals ||
		proofing.autoCorrectCapitalizeDayNames ||
		proofing.autoCorrectFractions ||
		proofing.autoCorrectOrdinals;
	if (needsWordPass) {
		result = result
			.split(/(\s+)/)
			.map((part) => (/^\s+$/.test(part) ? part : transformWord(part, proofing)))
			.join('');
	}
	if (proofing.autoCorrectCapitalizeFirstLetter) {
		result = capitalizeSentences(result, lead);
	}
	return result;
}

/** Index range `[start, end)` of `after` that differs from `before`, grown to whole words. */
function editedRange(before: string, after: string): [number, number] | null {
	if (before === after) {
		return null;
	}
	const limit = Math.min(before.length, after.length);
	let prefix = 0;
	while (prefix < limit && before.charCodeAt(prefix) === after.charCodeAt(prefix)) {
		prefix += 1;
	}
	let suffix = 0;
	while (
		suffix < limit - prefix &&
		before.charCodeAt(before.length - 1 - suffix) === after.charCodeAt(after.length - 1 - suffix)
	) {
		suffix += 1;
	}
	let start = prefix;
	let end = after.length - suffix;
	if (end <= start) {
		return null;
	}
	if (!/\s/.test(after.charAt(start))) {
		while (start > 0 && !/\s/.test(after.charAt(start - 1))) {
			start -= 1;
		}
	}
	if (!/\s/.test(after.charAt(end - 1))) {
		while (end < after.length && !/\s/.test(after.charAt(end))) {
			end += 1;
		}
	}
	return [start, end];
}

/**
 * AutoCorrect only what was typed in this edit. `previous` is the text before the
 * edit and `text` the committed text; the range that differs (grown to whole
 * words, like the word being completed) is corrected with the text before it as
 * context, and everything else is left exactly as it was, so a URL or a
 * collaborator's words are never rewritten. Without `previous` the whole text is
 * treated as typed.
 */
export function applyAutoCorrectToEdit(
	text: string,
	previous: string | undefined,
	proofing: ViewerProofingOptions,
): string {
	if (previous === undefined) {
		return applyAutoCorrect(text, proofing);
	}
	const range = editedRange(previous, text);
	if (!range) {
		return text;
	}
	const [start, end] = range;
	return (
		text.slice(0, start) +
		applyAutoCorrect(text.slice(start, end), proofing, text.slice(0, start)) +
		text.slice(end)
	);
}
