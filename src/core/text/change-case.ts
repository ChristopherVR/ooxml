/**
 * Office's Change Case modes (Word, PowerPoint and Visio Home > Font > Change Case):
 * UPPERCASE, lowercase, Sentence case., Capitalize Each Word and tOGGLE cASE.
 */
export type TextCaseMode = 'upper' | 'lower' | 'sentence' | 'capitalize' | 'toggle';
export const TEXT_CASE_MODES: readonly TextCaseMode[] = Object.freeze([
	'sentence',
	'lower',
	'upper',
	'capitalize',
	'toggle',
]);

/** One changed span of the original text, in UTF-16 offsets. */
export interface TextCaseRange {
	start: number;
	end: number;
	text: string;
}

const letter = /\p{L}/u;
const wordCharacter = /[\p{L}\p{N}\p{M}'’]/u;
const sentenceEnd = /[.!?。！？]/u;
const space = /\s/u;

interface Point {
	start: number;
	end: number;
	text: string;
	out: string;
}

/**
 * Converts each code point on its own, with sentence and word context taken from the whole
 * text, so a caller can split the result at run boundaries without losing that context.
 */
function convert(text: string, mode: TextCaseMode): Point[] {
	if (!TEXT_CASE_MODES.includes(mode)) throw new TypeError(`Unknown text case mode: ${mode}`);
	const points: Point[] = [];
	let offset = 0,
		sentenceStart = true,
		pendingEnd = false,
		previous = '';
	for (const point of text) {
		let out = point;
		if (letter.test(point)) {
			const lower = point.toLowerCase(),
				upper = point.toUpperCase();
			if (mode === 'upper') out = upper;
			else if (mode === 'lower') out = lower;
			else if (mode === 'toggle') out = point === upper && point !== lower ? lower : upper;
			else if (mode === 'sentence') out = sentenceStart ? upper : lower;
			else out = previous && wordCharacter.test(previous) ? lower : upper;
			sentenceStart = false;
			pendingEnd = false;
		} else if (point === '\n') {
			sentenceStart = true;
			pendingEnd = false;
		} else if (sentenceEnd.test(point)) pendingEnd = true;
		else if (space.test(point)) {
			if (pendingEnd) sentenceStart = true;
			pendingEnd = false;
		} else if (/\p{N}/u.test(point)) {
			sentenceStart = false;
			pendingEnd = false;
		}
		points.push({ start: offset, end: offset + point.length, text: point, out });
		offset += point.length;
		previous = point;
	}
	return points;
}

/** The whole text in the requested case. */
export function changeTextCase(text: string, mode: TextCaseMode): string {
	return convert(text, mode)
		.map((point) => point.out)
		.join('');
}

/**
 * Only the spans that change, split at every offset in `boundaries` (run or marker starts) and
 * never crossing a line break, so formatting runs keep their own characters.
 */
export function changeTextCaseRanges(
	text: string,
	mode: TextCaseMode,
	boundaries: Iterable<number> = [],
): TextCaseRange[] {
	const breaks = new Set(boundaries);
	const ranges: TextCaseRange[] = [];
	let current: TextCaseRange | undefined;
	for (const point of convert(text, mode)) {
		if (point.out === point.text) {
			current = undefined;
			continue;
		}
		if (current && current.end === point.start && !breaks.has(point.start)) {
			current.end = point.end;
			current.text += point.out;
		} else {
			current = { start: point.start, end: point.end, text: point.out };
			ranges.push(current);
		}
	}
	return ranges;
}
