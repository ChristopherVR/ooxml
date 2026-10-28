import {
	getInlineEditorSelectionResult,
	readEditableText,
	remapTextToSegments,
	textStylePatch,
	toggleSelectionBullets,
} from 'ooxml-ui/pptx';
import type { InlineTextEditSnapshot, ListStyleUpdateOptions } from 'ooxml-ui/pptx';
import type { PptxElement, TextStyle } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';

/**
 * Paragraph-level patch builders for the Home tab's Paragraph group:
 * bullet/numbered list, indent, alignment, and line spacing.
 * Lists author semantic paragraph bullets; other formatting uses the element
 * `textStyle` base, matching `editor-format-mutations.ts`.
 */

/** Indent step (px) applied per increase/decrease-indent click. */
const INDENT_STEP_PX = 24;

function textStyleBase(el: PptxElement): TextStyle {
	return hasTextProperties(el) ? (el.textStyle ?? {}) : {};
}

/**
 * Toggle the paragraph list type between the given kind and `'none'`.
 * `options.listConversionScope` sets how far a bare caret reaches (the host
 * `listConversionScope` prop; unset keeps PowerPoint's single paragraph).
 */
export function toggleListTypePatch(
	el: PptxElement,
	kind: 'bullet' | 'numbered',
	snapshot?: InlineTextEditSnapshot,
	options?: ListStyleUpdateOptions,
): Partial<PptxElement> {
	if (!hasTextProperties(el)) {
		return {};
	}
	const surface =
		typeof document === 'undefined'
			? null
			: document.querySelector<HTMLElement>('[data-inline-editor]');
	const liveText = snapshot?.text ?? (surface ? readEditableText(surface) : undefined);
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

/** Increase or decrease the paragraph left margin by one indent step. */
export function adjustIndentPatch(el: PptxElement, delta: 1 | -1): Partial<PptxElement> {
	const base = textStyleBase(el);
	const next = Math.max(0, (base.paragraphMarginLeft ?? 0) + delta * INDENT_STEP_PX);
	return textStylePatch(el, { paragraphMarginLeft: next });
}

/** Set the paragraph alignment. */
export function setAlignPatch(el: PptxElement, align: TextStyle['align']): Partial<PptxElement> {
	return textStylePatch(el, { align });
}

/** Set the line spacing multiplier (e.g. 1.15 = 115%). */
export function setLineSpacingPatch(el: PptxElement, lineSpacing: number): Partial<PptxElement> {
	return textStylePatch(el, { lineSpacing });
}
