/**
 * East Asian line-start and line-end prohibition (kinsoku shori): characters that may not begin
 * a line (closing brackets, commas, full stops, iteration marks) and characters that may not end
 * one (opening brackets, currency signs that lead a number).
 *
 * The default sets are Word 16's, measured from its output (a fixed-pitch line one character too
 * short, each candidate placed where it would start or end the second line), one set per East
 * Asian language. They are the sets ECMA-376 refers to in `w:kinsoku` (17.3.1.16); a document
 * replaces one language's set with `w:noLineBreaksBefore` / `w:noLineBreaksAfter` (17.15.1.58-59),
 * and `w:strictFirstAndLastChars` selects the strict Japanese set. Measured behaviour, on a Word 16
 * install with no East Asian editing language enabled (`LanguageSettings.LanguagePreferredForEditing`
 * false for Japanese, both Chinese and Korean; one with them enabled may differ):
 *
 *  - The set follows the East Asian language of the character's run (`w:lang/@w:eastAsia`).
 *    A run with no East Asian language gets no prohibition at all, whatever the theme's
 *    `w:themeFontLang/@w:eastAsia` says.
 *  - Korean has an empty default set: Word breaks Korean text before `。` or `」`, after ideographs
 *    and after Hangul alike.
 *  - A custom set replaces that language's default set; other languages keep theirs.
 *  - Combining marks and zero-width controls never start a line anyway (they belong to the
 *    preceding grapheme), so they are not listed.
 */

/** A language whose kinsoku set differs (Word's four East Asian line-breaking languages). */
export type KinsokuLanguage = 'ja' | 'zh-Hans' | 'zh-Hant' | 'ko';

/** Characters that may not start a line (`noStart`) and may not end one (`noEnd`). */
export interface KinsokuSets {
	noStart: string;
	noEnd: string;
}

const JAPANESE: KinsokuSets = {
	noStart: '!%),.:;?]}¢°’”‰′″℃、。々〉》」』】〕゛゜ゝゞ・ヽヾ！％），．：；？］｝｡｣､･ﾞﾟ￠',
	noEnd: '$([\\{£¥‘“〈《「『【〔＄（［｛｢￡￥',
};

/**
 * Strict Japanese (`w:strictFirstAndLastChars`, measured the same way): the standard set plus
 * small kana and the prolonged sound mark. The line-end set and other languages do not change.
 */
const JAPANESE_STRICT_EXTRA = 'ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶーｧｨｩｪｫｬｭｮｯｰ';

const DEFAULTS: Record<KinsokuLanguage, KinsokuSets> = {
	ja: JAPANESE,
	'zh-Hans': {
		noStart:
			'!%),.:;>?]}¢¨°·―‖’”…‰′″›℃、。〃〉》」』】〕〗〞︶︺︾﹀﹄﹚﹜﹞！＂％＇），．：；？］｀｜｝～￠',
		noEnd: '$([{£¥·‘“〈《「『【〔〖〝﹙﹛﹝＄（．［｛￡￥',
	},
	'zh-Hant': {
		noStart:
			'!),.:;?]}¢·’”•‥…‧′、。〉》」』】〕〞︰︱︳︴︶︸︺︼︾﹀﹂﹄﹏﹐﹑﹒﹔﹕﹖﹗﹚﹜﹞！），．：；？］｜｝､',
		noEnd: '([{£¥‘“‵〈《「『【〔〝︵︷︹︻︽︿﹁﹃﹙﹛﹝（｛',
	},
	ko: { noStart: '', noEnd: '' },
};

/** The kinsoku language of a BCP 47 tag (`ja-JP`, `zh-CN`, `zh-TW`, `ko-KR`); undefined otherwise. */
export function kinsokuLanguage(tag: string | undefined): KinsokuLanguage | undefined {
	if (!tag) return undefined;
	const lower = tag.toLowerCase();
	const [primary] = lower.split(/[-_]/);
	if (primary === 'ja') return 'ja';
	if (primary === 'ko') return 'ko';
	if (primary !== 'zh') return undefined;
	return /-(hant|tw|hk|mo)(-|$)/.test(lower) ? 'zh-Hant' : 'zh-Hans';
}

/** A document's replacement set for one language (`w:noLineBreaksBefore` or `...After`). */
export interface KinsokuOverride {
	/** The `w:lang` tag the set applies to. */
	language: string;
	chars: string;
}

/** Document-wide kinsoku settings. */
export interface KinsokuSettings {
	/** `w:strictFirstAndLastChars`: the strict Japanese line-start set. */
	strict?: boolean;
	/** `w:noLineBreaksBefore`: characters that may not start a line. */
	noLineBreaksBefore?: KinsokuOverride;
	/** `w:noLineBreaksAfter`: characters that may not end a line. */
	noLineBreaksAfter?: KinsokuOverride;
}

/** The sets that apply to text in `language` under a document's settings. */
export function kinsokuSets(
	language: KinsokuLanguage,
	settings: KinsokuSettings = {},
): KinsokuSets {
	const base = DEFAULTS[language];
	const custom = (override: KinsokuOverride | undefined) =>
		override && kinsokuLanguage(override.language) === language ? override.chars : undefined;
	const strictExtra = settings.strict && language === 'ja' ? JAPANESE_STRICT_EXTRA : '';
	return {
		noStart: custom(settings.noLineBreaksBefore) ?? base.noStart + strictExtra,
		noEnd: custom(settings.noLineBreaksAfter) ?? base.noEnd,
	};
}

const first = (text: string) => String.fromCodePoint(text.codePointAt(0) ?? 0);

/** Whether `grapheme` (its base character) may not start a line. */
export function prohibitedAtLineStart(grapheme: string, sets: KinsokuSets): boolean {
	return grapheme.length > 0 && sets.noStart.includes(first(grapheme));
}

/** Whether `grapheme` (its base character) may not end a line. */
export function prohibitedAtLineEnd(grapheme: string, sets: KinsokuSets): boolean {
	return grapheme.length > 0 && sets.noEnd.includes(first(grapheme));
}

/**
 * Line-end marks that `w:overflowPunct` (17.3.1.21) lets hang past the right margin instead of
 * pulling the character before them down, measured in Word 16. Word hangs them only in justified
 * (`both`, `distribute`) paragraphs with kinsoku on, one mark per line, and not when the next
 * character may not start a line either (`。」` wraps together). Each set was measured by placing
 * every line-start-prohibited character of the language where it would start the second line of
 * a justified paragraph: Japanese hangs only its commas, full stops and closing quotes and
 * brackets, Chinese most of its set. Korean has no prohibited characters, so nothing hangs.
 */
const HANGING: Record<KinsokuLanguage, string> = {
	ja: ',.’”、。」』】），．］｝｡､',
	'zh-Hans': '!%),.:;>?]}¢°’”‰′″℃、。〃〉》」』】〕〗〞﹚﹜﹞！＂％＇），：；？］｝￠',
	'zh-Hant': '!),.:;?]}’”′、。〉》」』】〕〞﹚﹜﹞！），．：；？］｝',
	ko: '',
};

/** Whether `grapheme` may hang at a line end under `w:overflowPunct`, in `language` text. */
export function hangsAtLineEnd(grapheme: string, language: KinsokuLanguage): boolean {
	return grapheme.length > 0 && HANGING[language].includes(first(grapheme));
}
