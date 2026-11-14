import { attachInlineRichPaste } from 'ooxml-ui/pptx';
import type { InlineRichPasteOptions } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { untrack } from 'svelte';

export interface InlineRichTextPasteDeps {
	editingElement: () => PptxElement | undefined;
	closeInline: () => void;
	enterInlineEdit: (id: string) => void;
	applyElementPatch: (id: string, patch: Partial<PptxElement>) => void;
}

/** The binding half of the shared rich text paste: close, write one history step, reopen. */
export function inlineRichTextPasteOptions(deps: InlineRichTextPasteDeps): InlineRichPasteOptions {
	return {
		getEditingElement: () => deps.editingElement() ?? null,
		cancelEdit: deps.closeInline,
		updateElement: (id, patch) => deps.applyElementPatch(id, patch as Partial<PptxElement>),
		enterEdit: deps.enterInlineEdit,
	};
}

/** Paste in the inline text editor keeps formatting from text/html (else pastes plain text). */
export function useInlineRichTextPaste(
	getRoot: () => HTMLElement | null,
	deps: InlineRichTextPasteDeps,
): void {
	const options = inlineRichTextPasteOptions(deps);
	$effect(() => {
		const root = getRoot();
		if (root) {
			return untrack(() => attachInlineRichPaste(root, options));
		}
	});
}
