import { attachInlineRichPaste } from 'ooxml-ui/pptx';
import type { InlineRichPasteOptions } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';
import { findActiveElement } from './editor-active-elements';

export interface InlineRichPasteDeps {
	store: Store<ViewerState>;
	inlineElementId(): string | null;
	closeInline(commit: boolean): void;
	openInlineEdit(id: string): void;
	pushHistory(): void;
	previewElement(id: string, patch: Partial<PptxElement>): void;
	commitChange(): void;
}

/** The binding half of the shared rich text paste: close, write one history step, reopen. */
export function inlineRichPasteOptions(deps: InlineRichPasteDeps): InlineRichPasteOptions {
	return {
		getEditingElement: () => {
			const id = deps.inlineElementId();
			return id ? (findActiveElement(deps.store.get(), id) ?? null) : null;
		},
		cancelEdit: () => {
			if (deps.inlineElementId() !== null) {
				deps.closeInline(false);
			}
		},
		updateElement: (id, patch) => {
			deps.pushHistory();
			deps.previewElement(id, patch as Partial<PptxElement>);
			deps.commitChange();
		},
		enterEdit: (id) => deps.openInlineEdit(id),
	};
}

/** Paste in the inline text editor keeps formatting from text/html (else pastes plain text). */
export function attachCanvasInlineRichPaste(
	root: HTMLElement,
	deps: InlineRichPasteDeps,
): () => void {
	return attachInlineRichPaste(root, inlineRichPasteOptions(deps));
}
