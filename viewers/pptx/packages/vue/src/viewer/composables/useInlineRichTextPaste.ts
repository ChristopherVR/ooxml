import { attachInlineRichPaste } from 'ooxml-ui/pptx';
import type { InlineRichPasteOptions } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import type { Ref } from 'vue';
import { watch } from 'vue';

export interface InlineRichTextPasteDeps {
	editingElement: () => PptxElement | undefined;
	cancelInlineEdit: () => void;
	enterInlineEdit: (id: string) => void;
	updateElement: (id: string, patch: Partial<PptxElement>) => void;
}

/** The binding half of the shared rich text paste: close, write one history step, reopen. */
export function inlineRichTextPasteOptions(deps: InlineRichTextPasteDeps): InlineRichPasteOptions {
	return {
		getEditingElement: () => deps.editingElement() ?? null,
		cancelEdit: deps.cancelInlineEdit,
		updateElement: (id, patch) => deps.updateElement(id, patch as Partial<PptxElement>),
		enterEdit: deps.enterInlineEdit,
	};
}

/** Paste in the inline text editor keeps formatting from text/html (else pastes plain text). */
export function useInlineRichTextPaste(
	root: Ref<HTMLElement | null>,
	deps: InlineRichTextPasteDeps,
): void {
	const options = inlineRichTextPasteOptions(deps);
	watch(
		root,
		(element, _previous, onCleanup) => {
			if (element) {
				onCleanup(attachInlineRichPaste(element, options));
			}
		},
		{ immediate: true, flush: 'sync' },
	);
}
