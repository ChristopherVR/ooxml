import type { InlineRichPasteOptions } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';
import { useMemo, useRef } from 'react';

export interface InlineRichTextPasteInput {
	elementLookup: Map<string, PptxElement>;
	inlineEditingElementIdRef: { current: string | null };
	setInlineEditingElementId: (id: string | null) => void;
	setInlineEditingText: (text: string) => void;
	updateElementById: (elementId: string, updates: Partial<PptxElement>) => void;
}

/**
 * The binding half of rich text paste into the inline editor (`attachInlineRichPaste`, shared):
 * make sure the editor is closed, write the pasted body as one history step, reopen it. `get`
 * returns the latest render's values.
 */
export function createInlineRichTextPasteOptions(
	get: () => InlineRichTextPasteInput,
): InlineRichPasteOptions {
	return {
		getEditingElement: () => {
			const id = get().inlineEditingElementIdRef.current;
			return id ? (get().elementLookup.get(id) ?? null) : null;
		},
		cancelEdit: () => {
			if (get().inlineEditingElementIdRef.current !== null) {
				get().setInlineEditingElementId(null);
				get().setInlineEditingText('');
			}
		},
		updateElement: (elementId, patch) =>
			get().updateElementById(elementId, patch as Partial<PptxElement>),
		enterEdit: (elementId) => {
			const element = get().elementLookup.get(elementId);
			if (!element || !hasTextProperties(element)) {
				return;
			}
			get().setInlineEditingElementId(elementId);
			get().setInlineEditingText(element.text ?? '');
		},
	};
}

export function useInlineRichTextPaste(input: InlineRichTextPasteInput): InlineRichPasteOptions {
	const latest = useRef(input);
	latest.current = input;
	return useMemo(() => createInlineRichTextPasteOptions(() => latest.current), []);
}
