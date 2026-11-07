/**
 * Mark the paragraph boundary Chromium inserts when Enter splits a seeded
 * rich run, so `readEditableText` reads a new paragraph there (an appended
 * numbered paragraph continues its list, #211) and native undo that rejoins
 * the split leaves no stray break. Every DOM-seeded plain editor (vanilla,
 * Svelte, Vue, Angular) calls it on `insertParagraph` input. Moved from the
 * vanilla binding.
 *
 * @module render/inline-text-paragraph-marker
 */
function isCaretPlaceholder(element: HTMLElement): boolean {
	const content = Array.from(element.childNodes).filter((child) => {
		if (child.nodeType === 3) {
			return (child.nodeValue ?? '').length > 0;
		}
		return child instanceof HTMLElement && !child.hasAttribute('data-pptx-bullet-marker');
	});
	if (content.length !== 1 || !(content[0] instanceof HTMLElement)) {
		return false;
	}
	return content[0].tagName === 'BR' || isCaretPlaceholder(content[0]);
}

function firstEditableRunIndex(element: HTMLElement): string | undefined {
	const run =
		element.matches('[data-seg-idx]') && !element.hasAttribute('data-pptx-bullet-marker')
			? element
			: element.querySelector<HTMLElement>('[data-seg-idx]:not([data-pptx-bullet-marker])');
	return run?.dataset.segIdx;
}

/** Mark Chromium's inserted rich-run paragraph boundary. */
export function markInsertedParagraph(doc: Document, surface: HTMLElement): void {
	const selection = doc.getSelection();
	const anchor = selection?.anchorNode;
	const anchorElement =
		anchor?.nodeType === Node.ELEMENT_NODE ? (anchor as HTMLElement) : anchor?.parentElement;
	const insertedRun = anchorElement?.closest<HTMLElement>('[data-seg-idx]');
	if (!insertedRun || !surface.contains(insertedRun)) {
		return;
	}
	const block = insertedRun.closest<HTMLElement>('div, p');
	const isTextFlow = block?.hasAttribute('data-pptx-text-flow');
	// A split flow owns its boundary. Marking a surviving authored run would
	// leave an extra break after native undo rejoins it into the original flow.
	const isSplitTextFlow =
		isTextFlow &&
		block?.previousElementSibling?.hasAttribute('data-pptx-text-flow') &&
		block.querySelector('[data-seg-idx]:not([data-pptx-bullet-marker])') === insertedRun;
	const caretParagraph =
		block && block !== surface && (!isTextFlow || isSplitTextFlow) && surface.contains(block)
			? block
			: insertedRun;
	const previous = caretParagraph.previousElementSibling;
	const paragraph =
		previous instanceof HTMLElement &&
		firstEditableRunIndex(previous) === insertedRun.dataset.segIdx &&
		isCaretPlaceholder(previous)
			? previous
			: caretParagraph;
	paragraph.dataset.pptxParagraphStart = '';
}
