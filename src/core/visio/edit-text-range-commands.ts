import { fail } from './package-common';

export interface VisioTextRange {
	/** Zero-based UTF16 scalar boundaries in the complete original logical text. */
	start: number;
	end: number;
	text: string;
}
export interface VisioTextRangesEdit {
	type: 'replace-text-ranges';
	pageId: string;
	shapeId: string;
	/** Complete logical text captured with the ranges; stale source offsets are rejected. */
	expectedText: string;
	/** Ascending, nonempty selected spans inside paragraphs; replacement text may be empty.
	 * Fields, paragraph break changes, competing character markers and empty rich paragraphs refuse.
	 */
	ranges: readonly VisioTextRange[];
}
const invalidXmlText =
	/[\u0000-\u0008\u000b\u000c\u000d\u000e-\u001f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/u;
export function textRangeBoundary(text: string, offset: number): boolean {
	return !(
		offset > 0 &&
		offset < text.length &&
		/[\ud800-\udbff]/.test(text[offset - 1]!) &&
		/[\udc00-\udfff]/.test(text[offset]!)
	);
}
/** Snapshot before any await; commands never retain caller-owned range arrays or objects. */
export function snapshotTextRanges(
	edit: VisioTextRangesEdit,
	consume?: (value: unknown) => string,
): VisioTextRangesEdit {
	let characters = 0;
	const text = (value: unknown): string => {
		if (typeof value !== 'string' || invalidXmlText.test(value))
			fail('INVALID_EDIT_TEXT', 'Range replacement requires valid XML text.');
		characters += value.length;
		if (characters > 1_000_000)
			fail('LIMIT_EDIT_TEXT', 'Range replacement text exceeds input limits.');
		return consume ? consume(value) : value;
	};
	const pageId = edit.pageId,
		shapeId = edit.shapeId;
	if ([pageId, shapeId].some((id) => typeof id !== 'string' || !id || id.length > 256))
		fail('INVALID_EDIT', 'Invalid text range target.');
	const expectedText = text(edit.expectedText),
		original = edit.ranges;
	if (!Array.isArray(original)) fail('INVALID_EDIT', 'Text ranges require an array.');
	const length = original.length;
	if (!Number.isSafeInteger(length) || length < 1 || length > 10000)
		fail('LIMIT_EDITS', 'Range replacement requires 1-10000 ranges.');
	const ranges: VisioTextRange[] = [];
	let end = 0,
		output = expectedText.length;
	for (let index = 0; index < length; index++) {
		const input = original[index]!,
			start = input.start,
			nextEnd = input.end,
			replacement = text(input.text);
		if (
			!Number.isSafeInteger(start) ||
			!Number.isSafeInteger(nextEnd) ||
			start < end ||
			nextEnd <= start ||
			nextEnd > expectedText.length ||
			!textRangeBoundary(expectedText, start) ||
			!textRangeBoundary(expectedText, nextEnd)
		)
			fail('INVALID_EDIT', 'Text ranges require ascending nonoverlapping UTF16 scalar boundaries.');
		if (replacement.includes('\n') || expectedText.slice(start, nextEnd).includes('\n'))
			fail('UNSUPPORTED_TEXT_RANGE', 'Range replacement across paragraph breaks is unsupported.');
		ranges.push({ start, end: nextEnd, text: replacement });
		end = nextEnd;
		output += replacement.length - (nextEnd - start);
	}
	if (output > 1_000_000) fail('LIMIT_EDIT_TEXT', 'Range replacement output exceeds limits.');
	return { type: 'replace-text-ranges', pageId, shapeId, expectedText, ranges };
}
