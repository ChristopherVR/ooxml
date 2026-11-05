/**
 * The element style of a new text element built from rich segments.
 *
 * @module sdk/text-segment-style
 */

import type { TextSegment, TextStyle } from '../../types/text';

/** Run-level keys the save path takes from the element style when every segment agrees. */
const RUN_KEYS = [
	'fontFamily',
	'fontSize',
	'bold',
	'italic',
	'underline',
	'strikethrough',
	'color',
	'colorRef',
	'characterSpacing',
	'hyperlink',
	'hyperlinkAction',
	'hyperlinkTargetSlideIndex',
	'hyperlinkTooltip',
	'align',
] as const satisfies ReadonlyArray<keyof TextStyle>;

/**
 * The element style for rich text built from segments, so the runs save as
 * given.
 *
 * When every segment has the same style, the save path writes the runs from
 * the element style alone, and for a key that every segment shares it lets
 * the element style override the segments. Built from the segments' common
 * values, the element style says what the segments say: a single bold
 * segment stays bold, and a key the segments agree on is not replaced.
 */
export function elementStyleForSegments(
	textStyle: Partial<TextStyle>,
	segments: readonly TextSegment[],
): Partial<TextStyle> {
	const styled = segments.filter((segment) => !segment.isParagraphBreak);
	const first = styled[0]?.style;
	if (!first) {
		return textStyle;
	}
	const next: Partial<TextStyle> = { ...textStyle };
	for (const key of RUN_KEYS) {
		const value = first[key];
		if (!styled.every((segment) => sameValue(segment.style?.[key], value))) {
			continue;
		}
		if (value === undefined) {
			delete next[key];
		} else {
			(next as Record<string, unknown>)[key] = value;
		}
	}
	return next;
}

function sameValue(a: unknown, b: unknown): boolean {
	if (a === b) {
		return true;
	}
	if (a && b && typeof a === 'object' && typeof b === 'object') {
		return JSON.stringify(a) === JSON.stringify(b);
	}
	return false;
}
