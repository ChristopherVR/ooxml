import type { TextSegment } from '../../core/types';
import { isBulletMarkerSegment } from './bullet-toggle';

/** The boolean run decorations the ribbon toggles. */
export type TextDecorationFlag = 'bold' | 'italic' | 'underline' | 'strikethrough';

export const TEXT_DECORATION_FLAGS: readonly TextDecorationFlag[] = [
	'bold',
	'italic',
	'underline',
	'strikethrough',
];

/** A run that carries text: not a paragraph break, not a display-only bullet marker. */
export function isContentSegment(segment: TextSegment): boolean {
	return (
		!segment.isParagraphBreak &&
		segment.text !== '\n' &&
		segment.text.length > 0 &&
		!isBulletMarkerSegment(segment)
	);
}
