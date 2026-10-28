import {
	applyTextStyleUpdate,
	getInlineEditorSelectionResult,
	remapTextToSegments,
	toggleSelectionBullets,
} from 'ooxml-ui/pptx';
import type { InlineTextEditSnapshot, ListStyleUpdateOptions } from 'ooxml-ui/pptx';
import type { PptxElement, TextStyle } from 'pptx-viewer-core';

import { canFormatText, readTextFormatState } from './editor-format-mutations';
import { currentInlineEditorText } from './inline-text-editor';

/**
 * Paragraph-level formatting-patch builders for the vanilla editor
 * (list type, indent, alignment, line spacing). Split out of
 * `editor-format-mutations.ts` (character-level formatting) to keep both
 * files within the project's file-size budget. List commands use the active
 * paragraph selection or caret; the other controls retain whole-element scope.
 */

/** Indent step in px applied by the increase/decrease indent buttons. */
const INDENT_STEP_PX = 24;

function patchTextStyle(el: PptxElement, patch: Partial<TextStyle>): Partial<PptxElement> {
	if (!canFormatText(el)) {
		return {};
	}
	return applyTextStyleUpdate(el, patch, null).patch;
}

/**
 * Toggle the selected paragraphs' list type; without an inline selection, target the element.
 * `options.listConversionScope` sets how far a bare caret reaches (the host
 * `listConversionScope` option; unset keeps PowerPoint's single paragraph).
 */
export function toggleListType(
	el: PptxElement,
	kind: Exclude<TextStyle['listType'], 'none' | undefined>,
	snapshot?: InlineTextEditSnapshot,
	options?: ListStyleUpdateOptions,
): Partial<PptxElement> {
	if (!canFormatText(el)) {
		return {};
	}
	const liveText = snapshot?.text ?? currentInlineEditorText();
	const current =
		liveText === undefined
			? el
			: {
					...el,
					text: liveText,
					textSegments:
						snapshot?.textSegments ?? remapTextToSegments(liveText, el.textSegments, el.textStyle),
				};
	const result = getInlineEditorSelectionResult(current.textSegments, { preserveCaret: true });
	if (result.kind !== 'supported' || (result.snapshot && result.snapshot.elementId !== el.id)) {
		return {};
	}
	return {
		...(liveText === undefined ? {} : { text: liveText }),
		...toggleSelectionBullets(
			current,
			kind,
			result.selection,
			result.snapshot?.textSegments,
			options,
		).patch,
	};
}

/** Step the paragraph left margin (indent) by `deltaSteps` * {@link INDENT_STEP_PX}, clamped >= 0. */
export function adjustIndent(el: PptxElement, deltaSteps: number): Partial<PptxElement> {
	const current = readTextFormatState(el).paragraphMarginLeft;
	const next = Math.max(0, current + deltaSteps * INDENT_STEP_PX);
	return patchTextStyle(el, { paragraphMarginLeft: next });
}

/** Set the paragraph text alignment element-wide. */
export function setTextAlign(el: PptxElement, align: TextStyle['align']): Partial<PptxElement> {
	return patchTextStyle(el, { align });
}

/** Set the line-spacing multiplier element-wide. */
export function setLineSpacing(el: PptxElement, lineSpacing: number): Partial<PptxElement> {
	return patchTextStyle(el, { lineSpacing });
}

/** Set the text flow direction (horizontal / rotated / stacked) element-wide. */
export function setTextDirection(
	el: PptxElement,
	textDirection: TextStyle['textDirection'],
): Partial<PptxElement> {
	return patchTextStyle(el, { textDirection });
}

/** Set how many columns the element's text body flows into (>= 1). */
export function setColumnCount(el: PptxElement, columnCount: number): Partial<PptxElement> {
	return patchTextStyle(el, { columnCount: Math.max(1, Math.round(columnCount)) });
}
