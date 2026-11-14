import { DestroyRef, effect, inject } from '@angular/core';
import { attachInlineRichPaste, slideSpaceElement } from 'ooxml-ui/pptx';
import type { InlineRichPasteOptions } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

export interface InlineRichPasteHost {
	rootElement(): HTMLElement | undefined;
	activeSlide(): PptxSlide | undefined;
	activeSlideIndex(): number;
	editingId(): string | null;
	closeEdit(): void;
	enterEdit(id: string): void;
	updateElement(slideIndex: number, id: string, patch: Partial<PptxElement>): void;
}

/** The binding half of the shared rich text paste: close, write one history step, reopen. */
export function inlineRichPasteOptions(host: InlineRichPasteHost): InlineRichPasteOptions {
	return {
		getEditingElement: () => {
			const id = host.editingId();
			const elements = host.activeSlide()?.elements;
			return id && elements ? slideSpaceElement(elements, id) : null;
		},
		cancelEdit: () => {
			if (host.editingId() !== null) {
				host.closeEdit();
			}
		},
		updateElement: (id, patch) =>
			host.updateElement(host.activeSlideIndex(), id, patch as Partial<PptxElement>),
		enterEdit: (id) => host.enterEdit(id),
	};
}

/** Paste in the inline text editor keeps formatting from text/html (else pastes plain text). */
export function setupViewerInlineRichPaste(host: InlineRichPasteHost): void {
	const options = inlineRichPasteOptions(host);
	let dispose = () => {};
	let bound: HTMLElement | undefined;
	inject(DestroyRef).onDestroy(() => dispose());
	effect(() => {
		const root = host.rootElement();
		if (root === bound) {
			return;
		}
		dispose();
		dispose = root ? attachInlineRichPaste(root, options) : () => {};
		bound = root;
	});
}
