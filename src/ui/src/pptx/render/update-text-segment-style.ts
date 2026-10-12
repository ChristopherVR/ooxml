import type { TextSegment, TextStyle } from 'ooxml-core/pptx';

import { mergeTextStyleUpdate } from './font-family-update';

/** Keep a runless paragraph's future body formatting in the same style edit. */
export function updateTextSegmentStyle(
	segment: TextSegment,
	updates: Partial<TextStyle>,
	options: { updateBodyStyle?: boolean } = {},
): TextSegment {
	return {
		...segment,
		...(options.updateBodyStyle === false
			? {}
			: { style: mergeTextStyleUpdate(segment.style, updates) }),
		...(segment.paragraphInsertionStyle
			? { paragraphInsertionStyle: mergeTextStyleUpdate(segment.paragraphInsertionStyle, updates) }
			: {}),
	};
}
