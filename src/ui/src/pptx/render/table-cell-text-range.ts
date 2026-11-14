/**
 * table-cell-text-range.ts: text-offset edits on rich segments for the table
 * cell editor. Offsets count a paragraph or line break as one character, as
 * `segmentsPlainText` spells them.
 *
 * @module render/table-cell-text-range
 */
import type { TextSegment, TextStyle } from 'ooxml-core/pptx';

import { updateTextSegmentStyle } from './update-text-segment-style';

function lengthOf(segment: TextSegment): number {
	return segment.isParagraphBreak || segment.isLineBreak ? 1 : segment.text.length;
}

/**
 * Apply `updates` to the characters in `[start, end)`, splitting the runs at
 * the range's edges. Breaks pass through unchanged.
 */
export function styleSegmentsRange(
	segments: readonly TextSegment[],
	start: number,
	end: number,
	updates: Partial<TextStyle>,
): TextSegment[] {
	const out: TextSegment[] = [];
	let offset = 0;
	for (const segment of segments) {
		const length = lengthOf(segment);
		const from = offset;
		offset += length;
		if (segment.isParagraphBreak || segment.isLineBreak || from >= end || offset <= start) {
			out.push(segment);
			continue;
		}
		const cutStart = Math.max(0, start - from);
		const cutEnd = Math.min(length, end - from);
		const before = segment.text.slice(0, cutStart);
		const middle = segment.text.slice(cutStart, cutEnd);
		const after = segment.text.slice(cutEnd);
		const { paragraphProperties, tableParagraph, ...rest } = segment;
		const pieces: TextSegment[] = [];
		if (before) {
			pieces.push({ ...rest, text: before, style: { ...segment.style } });
		}
		pieces.push(
			updateTextSegmentStyle({ ...rest, text: middle, style: { ...segment.style } }, updates),
		);
		if (after) {
			pieces.push({ ...rest, text: after, style: { ...segment.style } });
		}
		if (paragraphProperties) {
			pieces[0]!.paragraphProperties = paragraphProperties;
		}
		if (tableParagraph) {
			pieces[0]!.tableParagraph = tableParagraph;
		}
		out.push(...pieces);
	}
	return out;
}

/** The word around `offset` (non-space characters on either side), as PowerPoint formats it. */
export function wordRangeAt(text: string, offset: number): { start: number; end: number } {
	const isWord = (ch: string | undefined) => ch !== undefined && !/\s/.test(ch);
	let start = offset;
	let end = offset;
	while (isWord(text[start - 1])) {
		start--;
	}
	while (isWord(text[end])) {
		end++;
	}
	return { start, end };
}

/**
 * Turn the paragraph break at text offset `offset` into a soft line break
 * (`a:br`): the two paragraphs it separated become one, laid out as the first.
 */
export function toLineBreakAt(segments: readonly TextSegment[], offset: number): TextSegment[] {
	const out = segments.map((segment) => ({ ...segment }));
	let position = 0;
	for (let index = 0; index < out.length; index++) {
		const segment = out[index]!;
		if (position === offset && segment.isParagraphBreak) {
			const {
				isParagraphBreak: _p,
				paragraphProperties: _l,
				tableParagraph: _x,
				...rest
			} = segment;
			out[index] = { ...rest, text: '\n', isLineBreak: true };
			const next = out[index + 1];
			if (next) {
				delete next.paragraphProperties;
				delete next.tableParagraph;
			}
			return out;
		}
		position += lengthOf(segment);
		if (position > offset) {
			break;
		}
	}
	return out;
}
