/**
 * Offline English/Latin syllable hyphenator and soft-hyphen helpers.
 * Zero external dependencies or network requests; used by Word layout when
 * `w:autoHyphenation` is enabled or explicit soft hyphens (`\u00AD`) appear.
 */

export const SOFT_HYPHEN = '\u00AD';

export interface HyphenationOptions {
	/** Skip words composed entirely of uppercase letters (`w:doNotHyphenateCaps`). */
	doNotHyphenateCaps?: boolean;
	/** Minimum letters before the first break (default 2). */
	leftMin?: number;
	/** Minimum letters after the last break (default 3). */
	rightMin?: number;
}

const LATIN_LETTER = /\p{Script=Latin}/u;
const PURE_LATIN = /^\p{Script=Latin}+$/u;
const VOWEL = /[aeiouyàáâãäåæèéêëìíîïòóôõöøùúûüýÿ]/;

const VOWEL_DIGRAPHS = new Set('ai au ea ee ei eu ie oa oe oi oo ou ue ui'.split(' '));
const ONSETS_3 = new Set('str spr scr spl thr shr sch'.split(' '));
const ONSETS_2 = new Set(
	'bl br ch cl cr dr fl fr gh gl gr ph pl pr qu sh th tr tw wh wr'.split(' '),
);

const PREFIXES = (
	'counter contra pseudo extra hyper inter intra macro micro multi super trans ultra under ' +
	'anti auto fore meta mini mono omni over para poly post semi tele ' +
	'com con dis mal mid mis neo non out per pre pro sub tri uni ' +
	'bi co de en ex il im in ir re un'
).split(' ');

const SUFFIXES = (
	'ation ition ution cious tious cial tial sion tion ment ness less able ible ical ' +
	'ship ward wise like some teen ful ity ive ize ise ing ers est ous ary ery ory ish ist ism'
).split(' ');

/** Removes invisible soft hyphens (`\u00AD`) from text for measuring or rendering unbroken runs. */
export function stripSoftHyphens(text: string): string {
	return text.includes(SOFT_HYPHEN) ? text.replaceAll(SOFT_HYPHEN, '') : text;
}

/**
 * Removes U+FE0F (variation selector 16, "emoji presentation") from text for measuring or
 * writing it: Word gives it no width and draws nothing for it, even where the font maps it to a
 * glyph with an advance (Segoe UI Emoji: 2812 of 2048 units, #439).
 */
export function stripEmojiSelectors(text: string): string {
	return text.includes('️') ? text.replaceAll('️', '') : text;
}

/** Returns indices of explicit soft hyphens (`\u00AD`) that have text on both sides. */
export function softHyphenIndices(word: string): number[] {
	if (!word.includes(SOFT_HYPHEN)) return [];
	const indices: number[] = [];
	for (let i = 1; i < word.length - 1; i++) {
		if (word[i] === SOFT_HYPHEN) indices.push(i);
	}
	return indices;
}

/** Whether a word has alphabetic letters and every letter is uppercase. */
export function isAllCapsWord(word: string): boolean {
	return word === word.toUpperCase() && word !== word.toLowerCase();
}

function isVowelAt(word: string, index: number): boolean {
	const ch = word[index];
	if (!ch || !VOWEL.test(ch)) return false;
	// Leading 'y' or 'y' directly before a vowel acts as a consonant (e.g. "yellow", "beyond").
	if (
		ch === 'y' &&
		(index === 0 || (word[index + 1] !== undefined && VOWEL.test(word[index + 1]!)))
	)
		return false;
	// 'u' in 'qu' followed by another vowel is part of the consonant cluster (e.g. "require").
	if (
		ch === 'u' &&
		word[index - 1] === 'q' &&
		word[index + 1] !== undefined &&
		VOWEL.test(word[index + 1]!)
	)
		return false;
	return true;
}

function hasVowel(slice: string): boolean {
	for (let i = 0; i < slice.length; i++) {
		if (isVowelAt(slice, i)) return true;
	}
	return false;
}

interface Span {
	start: number;
	end: number;
}

function vowelNuclei(word: string): Span[] {
	const nuclei: Span[] = [];
	let i = 0;
	while (i < word.length) {
		if (!isVowelAt(word, i)) {
			i++;
			continue;
		}
		const start = i;
		i++;
		while (i < word.length && isVowelAt(word, i)) {
			const pair = word.slice(i - 1, i + 1);
			if (!VOWEL_DIGRAPHS.has(pair)) break;
			i++;
		}
		nuclei.push({ start, end: i });
	}
	// A trailing silent 'e' is not its own syllable unless preceded by consonant + 'l' (e.g. "table").
	if (nuclei.length > 1) {
		const last = nuclei.at(-1)!;
		if (last.start === word.length - 1 && word[last.start] === 'e') {
			const syllabicLe =
				word[last.start - 1] === 'l' && last.start >= 2 && !isVowelAt(word, last.start - 2);
			if (!syllabicLe) nuclei.pop();
		}
	}
	return nuclei;
}

function splitConsonantCluster(word: string, start: number, end: number): number {
	const count = end - start;
	if (count === 0) return start;
	if (count === 1) return word[start] === 'x' ? end : start;
	if (count === 2) {
		const pair = word.slice(start, end);
		if (pair[0] === pair[1]) return start + 1;
		if (pair === 'ck' || pair === 'ng') return end;
		if (ONSETS_2.has(pair)) return start;
		return start + 1;
	}
	const tail3 = word.slice(end - 3, end);
	if (ONSETS_3.has(tail3) && end - 3 > start) return end - 3;
	const tail2 = word.slice(end - 2, end);
	if (ONSETS_2.has(tail2) && end - 2 > start) return end - 2;
	return start + 1;
}

function isolateLatinCore(word: string): { core: string; offset: number } | undefined {
	let start = 0;
	let end = word.length;
	while (start < end && !LATIN_LETTER.test(word[start]!)) start++;
	while (end > start && !LATIN_LETTER.test(word[end - 1]!)) end--;
	const core = word.slice(start, end);
	if (!core || !PURE_LATIN.test(core)) return undefined;
	return { core, offset: start };
}

/**
 * Returns sorted character indices within `word` where an automatic hyphen may be inserted.
 * Each index `i` splits `word` into `word.slice(0, i) + '-'` and `word.slice(i)`.
 */
export function hyphenationPoints(word: string, options?: HyphenationOptions): number[] {
	const leftMin = options?.leftMin ?? 2;
	const rightMin = options?.rightMin ?? 3;
	const isolated = isolateLatinCore(word);
	if (!isolated) return [];
	const { core, offset } = isolated;
	if (core.length < leftMin + rightMin) return [];
	if (options?.doNotHyphenateCaps && isAllCapsWord(core)) return [];

	const lower = core.toLowerCase();
	const breaks = new Set<number>();

	for (const prefix of PREFIXES) {
		if (lower.startsWith(prefix) && lower.length - prefix.length >= rightMin) {
			const rest = lower.slice(prefix.length);
			const boundaryPair = lower.slice(prefix.length - 1, prefix.length + 1);
			if (hasVowel(rest) && !VOWEL_DIGRAPHS.has(boundaryPair)) {
				breaks.add(prefix.length);
				break;
			}
		}
	}

	for (const suffix of SUFFIXES) {
		if (lower.endsWith(suffix) && lower.length - suffix.length >= leftMin) {
			const stemEnd = lower.length - suffix.length;
			const stem = lower.slice(0, stemEnd);
			// Doubled consonant before vowel-initial suffix ("run-ning", "for-mat-ting").
			if (
				stemEnd >= 2 &&
				isVowelAt(lower, stemEnd) &&
				lower[stemEnd - 1] === lower[stemEnd - 2] &&
				!isVowelAt(lower, stemEnd - 1)
			) {
				if (hasVowel(lower.slice(0, stemEnd - 1))) breaks.add(stemEnd - 1);
			} else if (hasVowel(stem)) {
				breaks.add(stemEnd);
			}
			break;
		}
	}

	const nuclei = vowelNuclei(lower);
	for (let i = 0; i + 1 < nuclei.length; i++) {
		const prev = nuclei[i]!;
		const next = nuclei[i + 1]!;
		// Skip if a prefix or suffix break already divides this inter-vocalic consonant span.
		let covered = false;
		for (const existing of breaks) {
			if (existing >= prev.end && existing <= next.start) {
				covered = true;
				break;
			}
		}
		if (!covered) breaks.add(splitConsonantCluster(lower, prev.end, next.start));
	}

	const maxPos = core.length - rightMin;
	return [...breaks]
		.filter((pos) => pos >= leftMin && pos <= maxPos)
		.sort((a, b) => a - b)
		.map((pos) => pos + offset);
}
