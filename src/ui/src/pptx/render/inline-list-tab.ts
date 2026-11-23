import type { TextSegment } from 'ooxml-core/pptx';
import { breakAutoNumberRun, createAutoNumberSequence, nextAutoNumber } from 'ooxml-core/pptx';

import { formatAutoNumber } from './bullet-autonum';
import { resolveParagraphBullet } from './bullet-list';
import { isBulletMarkerSegment, paragraphBulletKind, splitBulletParagraphs } from './bullet-toggle';
import type { BulletParagraph } from './bullet-toggle';

/** Highest OOXML outline/list level (`a:pPr/@lvl` accepts `0..8`). */
export const MAX_BULLET_LIST_LEVEL = 8;

export type InlineListTabDirection = 'demote' | 'promote';

export type InlineListTabIndentResult =
	| { handled: false }
	| { handled: true; changed: boolean; textSegments: TextSegment[] };

interface ParagraphBodySpan {
	start: number;
	end: number;
}

function paragraphBodyLength(paragraph: BulletParagraph): number {
	let length = 0;
	for (const [index, segment] of paragraph.segments.entries()) {
		if (index === 0 && isBulletMarkerSegment(segment)) {
			continue;
		}
		length += segment.isLineBreak ? 1 : segment.text.length;
	}
	return length;
}

function paragraphBodySpans(paragraphs: readonly BulletParagraph[]): ParagraphBodySpan[] {
	const spans: ParagraphBodySpan[] = [];
	let cursor = 0;
	for (const [index, paragraph] of paragraphs.entries()) {
		if (index > 0) {
			cursor += 1; // '\n' paragraph break
		}
		const length = paragraphBodyLength(paragraph);
		spans.push({ start: cursor, end: cursor + length });
		cursor += length;
	}
	return spans;
}

function paragraphIndexForOffset(spans: readonly ParagraphBodySpan[], offset: number): number {
	for (const [index, span] of spans.entries()) {
		if (offset <= span.end) {
			return index;
		}
	}
	return Math.max(0, spans.length - 1);
}

function clampListLevel(level: number | undefined): number {
	if (typeof level !== 'number' || !Number.isFinite(level)) {
		return 0;
	}
	return Math.min(MAX_BULLET_LIST_LEVEL, Math.max(0, Math.trunc(level)));
}

function recomputeAutoNumberOrdinals(paragraphs: readonly BulletParagraph[]): void {
	const sequence = createAutoNumberSequence();
	for (const paragraph of paragraphs) {
		const first = paragraph.segments[0];
		if (!first) {
			continue;
		}
		const level = clampListLevel(first.paragraphLevel);
		const bullet = resolveParagraphBullet(first);
		if (bullet?.isNumbered && first.bulletInfo?.autoNumType) {
			const wasMarker = isBulletMarkerSegment(first);
			const start = first.bulletInfo.autoNumStartAt ?? 1;
			const nextIndex =
				nextAutoNumber(sequence, level, first.bulletInfo.autoNumType, start) - start;
			first.bulletInfo = { ...first.bulletInfo, paragraphIndex: nextIndex };
			if (wasMarker) {
				const trailingSpace = first.text.endsWith(' ') ? ' ' : '';
				first.text =
					formatAutoNumber(first.bulletInfo.autoNumType, Math.max(1, start + nextIndex)) +
					trailingSpace;
			}
		} else {
			breakAutoNumberRun(sequence, level);
		}
	}
}

/**
 * Resolve Tab (`demote`) and Shift+Tab (`promote`) inside an inline text editor.
 *
 * - `demote`: applies when the caret is at the start of a bulleted/numbered
 *   paragraph or when a multi-paragraph selection covers at least one
 *   bulleted/numbered paragraph. Levels clamp at {@link MAX_BULLET_LIST_LEVEL}.
 * - `promote`: applies when the caret or selection covers at least one
 *   bulleted/numbered paragraph (at any caret position). Levels clamp at `0`.
 * - Returns `{ handled: false }` when the key should fall back to normal
 *   behavior (inserting `\t` for Tab, or suppressing focus exit for Shift+Tab).
 */
export function applyInlineListTabIndent(
	segments: readonly TextSegment[],
	bodyRange: { start: number; end: number },
	direction: InlineListTabDirection,
): InlineListTabIndentResult {
	if (segments.length === 0) {
		return { handled: false };
	}
	const cloned = structuredClone(segments) as TextSegment[];
	const paragraphs = splitBulletParagraphs(cloned);
	const spans = paragraphBodySpans(paragraphs);
	if (spans.length === 0) {
		return { handled: false };
	}

	const rangeStart = Math.max(0, Math.min(bodyRange.start, bodyRange.end));
	const rangeEnd = Math.max(rangeStart, Math.max(bodyRange.start, bodyRange.end));
	const collapsed = rangeStart === rangeEnd;
	const startPara = paragraphIndexForOffset(spans, rangeStart);
	let endPara = collapsed ? startPara : paragraphIndexForOffset(spans, rangeEnd);
	const endSpan = spans[endPara];
	if (
		!collapsed &&
		endPara > startPara &&
		endSpan !== undefined &&
		rangeEnd === endSpan.start &&
		endSpan.end > endSpan.start
	) {
		endPara -= 1;
	}

	const startSpan = spans[startPara];
	if (direction === 'demote') {
		const atParagraphStart = collapsed && startSpan !== undefined && rangeStart === startSpan.start;
		const multiParagraph = endPara > startPara;
		if (!atParagraphStart && !multiParagraph) {
			return { handled: false };
		}
	}

	let hasBulletedTarget = false;
	let changed = false;
	const delta = direction === 'demote' ? 1 : -1;

	for (let index = startPara; index <= endPara; index += 1) {
		const paragraph = paragraphs[index];
		if (!paragraph || paragraph.segments.length === 0) {
			continue;
		}
		if (paragraphBulletKind(paragraph.segments) === 'none') {
			continue;
		}
		hasBulletedTarget = true;
		const first = paragraph.segments[0];
		if (!first) {
			continue;
		}
		const currentLevel = clampListLevel(first.paragraphLevel);
		const nextLevel = Math.min(MAX_BULLET_LIST_LEVEL, Math.max(0, currentLevel + delta));
		if (nextLevel !== currentLevel) {
			changed = true;
			first.paragraphLevel = nextLevel;
			for (let segIdx = 1; segIdx < paragraph.segments.length; segIdx += 1) {
				const seg = paragraph.segments[segIdx];
				if (seg && seg.paragraphLevel !== undefined) {
					seg.paragraphLevel = nextLevel;
				}
			}
		}
	}

	if (!hasBulletedTarget) {
		return { handled: false };
	}

	if (changed) {
		recomputeAutoNumberOrdinals(paragraphs);
	}

	return {
		handled: true,
		changed,
		textSegments: cloned,
	};
}
