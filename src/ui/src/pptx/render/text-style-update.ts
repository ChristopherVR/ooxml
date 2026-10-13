/**
 * The one text-style update every binding applies from the ribbon, the
 * inspector and the inline editor's keyboard shortcuts.
 *
 * With an inline selection the run-level half of `updates` is scoped to the
 * selected characters (`applyStyleToSelectedSegments`) while the body-level
 * half (`text-style-scope.ts`) still reaches `element.textStyle`; the
 * decoration flags the edit touched are then reconciled so the body's CSS
 * decoration agrees with the runs. Without a selection the whole update goes
 * to the body and to every run, the behaviour all five bindings already had.
 */

import type { PptxElement, TextSegment, TextStyle } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';
import { textRunsStylePatch } from 'ooxml-core/pptx/editor/render/text-runs-style-patch';

import { mergeTextStyleUpdate } from './font-family-update';
import type { InlineTextSelection } from './inline-selection-utils';
import { applyStyleToSelectedSegments } from './inline-selection-utils';
import type { TextDecorationFlag } from './selection-format-state';
import { TEXT_DECORATION_FLAGS } from './selection-format-state';
import {
	overrideParagraphIndents,
	overrideParagraphProperties,
	reconcileDecorationFlags,
	splitTextStyleUpdate,
	withExclusiveLineSpacing,
} from './text-style-scope';

/** What a text-style update hands back to the binding. */
export interface TextStyleUpdateResult {
	/** The element patch to apply through the binding's update-element operation. */
	patch: Partial<PptxElement>;
	/** The selection remapped onto the new segments; `null` when there was none. */
	newSelection: InlineTextSelection | null;
}

/**
 * Build the element patch for `updates`.
 *
 * @param selection - The active inline text selection, or `null` for a
 *   whole-element edit.
 * @param segments - A fresher copy of the element's segments (a live inline
 *   edit whose text is not yet on the model); defaults to the element's own.
 */
export function applyTextStyleUpdate(
	element: PptxElement,
	updates: Partial<TextStyle>,
	selection: InlineTextSelection | null,
	segments?: readonly TextSegment[],
): TextStyleUpdateResult {
	if (!hasTextProperties(element)) {
		return { patch: {}, newSelection: null };
	}
	const current = segments ?? element.textSegments;
	const exclusive = withExclusiveLineSpacing(updates);

	if (selection && current && current.length > 0) {
		const { body, run } = splitTextStyleUpdate(exclusive);
		let textStyle: TextStyle = mergeTextStyleUpdate(element.textStyle, body);
		let textSegments: TextSegment[] = overrideParagraphProperties(
			current,
			exclusive,
			element.textStyle?.inheritedParagraphSpacing,
		);
		let newSelection = selection;
		if (Object.keys(run).length > 0) {
			const scoped = applyStyleToSelectedSegments(textSegments, selection, run, element.textStyle);
			textSegments = scoped.newSegments;
			newSelection = scoped.newSelection;
		}
		const touched = TEXT_DECORATION_FLAGS.filter((flag): flag is TextDecorationFlag => flag in run);
		if (touched.length > 0) {
			const reconciled = reconcileDecorationFlags(textSegments, textStyle, touched);
			textStyle = reconciled.textStyle;
			textSegments = reconciled.textSegments;
		}
		return {
			patch: {
				textStyle,
				textSegments,
				...overrideParagraphIndents(element, exclusive),
			} as Partial<PptxElement>,
			newSelection,
		};
	}

	const wholeText = textRunsStylePatch(
		{
			...element,
			textSegments:
				current &&
				overrideParagraphProperties(
					current,
					exclusive,
					element.textStyle?.inheritedParagraphSpacing,
				),
		} as PptxElement,
		exclusive,
	);
	return {
		patch: {
			...wholeText,
			...overrideParagraphIndents(element, exclusive),
		} as Partial<PptxElement>,
		newSelection: null,
	};
}
