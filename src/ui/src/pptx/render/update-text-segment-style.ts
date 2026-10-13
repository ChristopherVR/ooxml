import type { TextSegment, TextStyle } from 'ooxml-core/pptx';

import { mergeTextStyleUpdate } from './font-family-update';

/**
 * Keep a runless paragraph's future body formatting in the same style edit.
 * `options.inherited` is the text box style the run inherits unset fields
 * from, so a font family pick can tell which script faces matched the old
 * effective Latin face (see `font-family-update.ts`).
 */
export function updateTextSegmentStyle(
	segment: TextSegment,
	updates: Partial<TextStyle>,
	options: { updateBodyStyle?: boolean; inherited?: TextStyle | undefined } = {},
): TextSegment {
	const { inherited } = options;
	return {
		...segment,
		...(options.updateBodyStyle === false
			? {}
			: { style: mergeTextStyleUpdate(segment.style, updates, inherited) }),
		...(segment.paragraphInsertionStyle
			? {
					paragraphInsertionStyle: mergeTextStyleUpdate(
						segment.paragraphInsertionStyle,
						updates,
						inherited,
					),
				}
			: {}),
	};
}
