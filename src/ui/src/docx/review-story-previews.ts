import type { Block, DocumentModel, ReviewDisplayMode } from 'ooxml-core/docx';
import { effectiveHeaderFooter } from './header-footer-link';
import type { HeaderFooterSlotName } from './header-footer-editor';
import { renderBlocks } from './header-footer-view';

/** Update inactive previews without rebuilding or closing an active story editor. */
export function refreshStoryPreviews(
	model: DocumentModel,
	sectionIndex: number,
	elements: {
		headers: HTMLElement | undefined;
		footers: HTMLElement | undefined;
		notes: HTMLElement | undefined;
	},
	mode: ReviewDisplayMode,
	decorate: (body: HTMLElement) => void,
): void {
	const refresh = (body: HTMLElement | null, blocks: Block[] | undefined) => {
		if (!body || !blocks || body.querySelector('.ProseMirror')) return;
		body.replaceChildren(renderBlocks(blocks, model, mode));
		decorate(body);
	};
	for (const kind of ['headers', 'footers'] as const)
		for (const slot of elements[kind]?.querySelectorAll<HTMLElement>('[data-slot]') ?? [])
			refresh(
				slot.querySelector('.dve-header-footer-body'),
				effectiveHeaderFooter(model, sectionIndex, kind, slot.dataset.slot as HeaderFooterSlotName)
					?.blocks,
			);
	const notes = [...(model.footnotes ?? []), ...(model.endnotes ?? [])];
	for (const item of elements.notes?.querySelectorAll<HTMLElement>('[data-docx-note-id]') ?? [])
		refresh(
			item.querySelector('.dve-note-body'),
			notes.find((note) => note.id === item.dataset.docxNoteId)?.blocks,
		);
}
