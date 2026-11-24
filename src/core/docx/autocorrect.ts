// Word AutoCorrect and AutoFormat As You Type rules (DOM-free): smart quotes, automatic bulleted
// and numbered lists, dashes, fractions, common symbols and common English typo replacements.

export interface AutoCorrectOptions {
	/** Replace straight quotes `"` and `'` with curly smart quotes as you type (default true). */
	smartQuotes?: boolean;
	/** Convert `- `, `* `, `1. ` and `1) ` at paragraph start into lists (default true). */
	autoLists?: boolean;
	/** Replace `--` with em dash `—` and ` - ` with en dash ` – ` (default true). */
	dashes?: boolean;
	/** Replace `1/2`, `1/4`, `3/4` with `½`, `¼`, `¾` on word boundary (default true). */
	fractions?: boolean;
	/** Replace `(c)`, `(r)`, `(tm)`, `...`, arrows (`-->`, `<--`, `<->`) (default true). */
	symbols?: boolean;
	/** Replace common English typos on word boundary (default true). */
	commonTypos?: boolean;
}

export type AutoCorrectAction =
	| {
			readonly kind: 'replace';
			/** Number of characters immediately before the insertion point to replace. */
			readonly deleteBefore: number;
			/** Replacement text (including any trailing trigger character such as space). */
			readonly text: string;
	  }
	| {
			readonly kind: 'list';
			readonly listKind: 'bullet' | 'decimal';
			/** Number of characters immediately before the insertion point to remove (`-`, `*`, `1.`). */
			readonly deleteBefore: number;
	  };

const FRACTIONS: Readonly<Record<string, string>> = {
	'1/2': '\u00BD',
	'1/4': '\u00BC',
	'3/4': '\u00BE',
};

const SYMBOLS: Readonly<Record<string, string>> = {
	'(c)': '\u00A9',
	'(C)': '\u00A9',
	'(r)': '\u00AE',
	'(R)': '\u00AE',
	'(tm)': '\u2122',
	'(TM)': '\u2122',
	'-->': '\u2192',
	'<--': '\u2190',
	'<->': '\u2194',
	'==>': '\u21D2',
	'<==': '\u21D0',
};

const COMMON_TYPOS: Readonly<Record<string, string>> = {
	adn: 'and',
	becuase: 'because',
	hte: 'the',
	occured: 'occurred',
	recieve: 'receive',
	seperate: 'separate',
	taht: 'that',
	teh: 'the',
	untill: 'until',
	waht: 'what',
	wiht: 'with',
};

const BOUNDARY_CHARS = new Set([' ', '\t', '.', ',', ';', ':', '!', '?', ')']);
const OPEN_QUOTE_PREV = /(?:^|[\s([{<\u2014\u2013])$/;

function matchCase(source: string, replacement: string): string {
	if (source === source.toUpperCase() && source !== source.toLowerCase())
		return replacement.toUpperCase();
	const first = source.charAt(0);
	if (first === first.toUpperCase() && first !== first.toLowerCase())
		return replacement.charAt(0).toUpperCase() + replacement.slice(1);
	return replacement;
}

/**
 * Evaluates AutoCorrect / AutoFormat As You Type for a single text insertion `inserted` at the
 * end of `beforeText` (the paragraph text immediately preceding the caret).
 *
 * Returns an action when a rule matches, or `undefined` to let normal typing proceed unchanged.
 */
export function autoCorrectTextInput(
	beforeText: string,
	inserted: string,
	options: AutoCorrectOptions = {},
): AutoCorrectAction | undefined {
	if (inserted.length !== 1) return undefined;
	const smartQuotes = options.smartQuotes ?? true;
	const autoLists = options.autoLists ?? true;
	const dashes = options.dashes ?? true;
	const fractions = options.fractions ?? true;
	const symbols = options.symbols ?? true;
	const commonTypos = options.commonTypos ?? true;

	// 1. Automatic bulleted / numbered lists at paragraph start on space
	if (autoLists && inserted === ' ') {
		if (beforeText === '-' || beforeText === '*' || beforeText === '\u2022') {
			return { kind: 'list', listKind: 'bullet', deleteBefore: beforeText.length };
		}
		if (beforeText === '1.' || beforeText === '1)') {
			return { kind: 'list', listKind: 'decimal', deleteBefore: beforeText.length };
		}
	}

	// 2. Smart quotes (`"` and `'`)
	if (smartQuotes) {
		if (inserted === '"') {
			const open = OPEN_QUOTE_PREV.test(beforeText);
			return { kind: 'replace', deleteBefore: 0, text: open ? '\u201C' : '\u201D' };
		}
		if (inserted === "'") {
			const open = OPEN_QUOTE_PREV.test(beforeText);
			return { kind: 'replace', deleteBefore: 0, text: open ? '\u2018' : '\u2019' };
		}
	}

	// 3. Immediate symbol & dash replacements (triggered as the final character is typed)
	const combined = beforeText + inserted;
	if (symbols) {
		if (combined.endsWith('...') && !beforeText.endsWith('...')) {
			return { kind: 'replace', deleteBefore: 2, text: '\u2026' };
		}
		for (const [trigger, symbol] of Object.entries(SYMBOLS)) {
			if (combined.endsWith(trigger)) {
				return { kind: 'replace', deleteBefore: trigger.length - 1, text: symbol };
			}
		}
	}

	if (dashes && inserted === '-' && beforeText.endsWith('-') && !beforeText.endsWith('--')) {
		return { kind: 'replace', deleteBefore: 1, text: '\u2014' };
	}

	// 4. Word-boundary rules (space or punctuation after a token)
	if (BOUNDARY_CHARS.has(inserted)) {
		// `word - ` -> `word – ` (en dash surrounded by spaces)
		if (dashes && inserted === ' ' && /\S -$/.test(beforeText)) {
			return { kind: 'replace', deleteBefore: 1, text: '\u2013 ' };
		}
		const tokenMatch = /([^\s]+)$/.exec(beforeText);
		const token = tokenMatch?.[1];
		if (token) {
			if (fractions && Object.prototype.hasOwnProperty.call(FRACTIONS, token)) {
				return {
					kind: 'replace',
					deleteBefore: token.length,
					text: FRACTIONS[token]! + inserted,
				};
			}
			if (commonTypos) {
				const lower = token.toLowerCase();
				if (Object.prototype.hasOwnProperty.call(COMMON_TYPOS, lower)) {
					return {
						kind: 'replace',
						deleteBefore: token.length,
						text: matchCase(token, COMMON_TYPOS[lower]!) + inserted,
					};
				}
			}
		}
	}

	return undefined;
}
