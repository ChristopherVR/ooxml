import type { PptxElement, PptxElementWithText, TextStyle } from '../../core/types';
import { hasTextProperties } from '../../core/types';
import { mergeTextStyleUpdate } from './font-family-update';
import { updateTextSegmentStyle } from './update-text-segment-style';

/** Immutable whole-text edit, using the pre-edit body style for run inheritance. */
export function textRunsStylePatch(
	element: PptxElement,
	updates: Partial<TextStyle>,
): Pick<Partial<PptxElementWithText>, 'textStyle' | 'textSegments'> {
	if (!hasTextProperties(element)) return {};
	const textStyle = mergeTextStyleUpdate(element.textStyle, updates);
	const textSegments = element.textSegments?.map((segment) =>
		updateTextSegmentStyle(segment, updates, { inherited: element.textStyle }),
	);
	return { textStyle, ...(textSegments ? { textSegments } : {}) };
}
