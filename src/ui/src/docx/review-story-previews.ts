import type { Block, DocumentModel, ReviewDisplayMode } from 'ooxml-core/docx';
import { effectiveHeaderFooter } from './header-footer-link';
import type { HeaderFooterSlotName } from './header-footer-editor';
import { renderStoryPreview } from './story-render';
import type { PictureSource } from './picture-paint';

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
	pictureUrl?: PictureSource,
): void {
	const refresh = (body: HTMLElement | null, blocks: Block[] | undefined) => {
		if (!body || !blocks || body.querySelector('.ProseMirror')) return;
		body.replaceChildren(renderStoryPreview(blocks, model, mode, pictureUrl));
	};
	for (const kind of ['headers', 'footers'] as const)
		for (const slot of elements[kind]?.querySelectorAll<HTMLElement>('[data-slot]') ?? [])
			refresh(
				slot.querySelector('.dve-header-footer-body'),
				effectiveHeaderFooter(model, sectionIndex, kind, slot.dataset.slot as HeaderFooterSlotName)
					?.blocks,
			);
	for (const item of elements.notes?.querySelectorAll<HTMLElement>('[data-docx-note-id]') ?? []) {
		const notes = item.dataset.docxNoteKind === 'endnote' ? model.endnotes : model.footnotes;
		refresh(
			item.querySelector('.dve-note-body'),
			notes?.find((note) => note.id === item.dataset.docxNoteId)?.blocks,
		);
	}
}
