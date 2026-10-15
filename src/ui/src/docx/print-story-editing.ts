import type { Block, DocumentModel } from 'ooxml-core/docx';
import type { PartsController } from './parts-controller';
import type { HeaderFooterSlotName } from './header-footer-editor';
import { effectiveHeaderFooter } from './header-footer-link';
import { newHeaderFooterId } from './header-footer-commands';
import { renderHeaderFooter } from './print-header-footer';

export interface PrintStoryHost {
	/** The `.dve-print-pages` element. */
	pages: HTMLElement;
	parts: Pick<PartsController, 'editOnPage' | 'closeOpenStory'>;
	model(): DocumentModel;
	printActive(): boolean;
	holdLayout(held: boolean): void;
}

const EDITING = 'dve-print-story-editing';

/** The header or footer area of a page under a viewport point, or undefined for the body. */
function areaAt(
	sheet: HTMLElement,
	clientY: number,
): { kind: 'headers' | 'footers'; element: HTMLElement | null } | undefined {
	const rect = sheet.getBoundingClientRect();
	const scale = (sheet.offsetHeight ? rect.height / sheet.offsetHeight : 1) || 1;
	const y = (clientY - rect.top) / scale;
	const header = sheet.querySelector<HTMLElement>(':scope > .dve-print-header');
	const footer = sheet.querySelector<HTMLElement>(':scope > .dve-print-footer');
	// A tall header or footer reaches into the body area; its own box counts as the story.
	const headerBottom = header ? (header.getBoundingClientRect().bottom - rect.top) / scale : 0;
	if (y < Math.max(Number(sheet.dataset.marginTop), headerBottom))
		return { kind: 'headers', element: header };
	const footerTop = footer ? (footer.getBoundingClientRect().top - rect.top) / scale : Infinity;
	if (y > Math.min(sheet.offsetHeight - Number(sheet.dataset.marginBottom), footerTop))
		return { kind: 'footers', element: footer };
	return undefined;
}

/**
 * Double-clicking a page's header or footer area edits it in place, as in Word: the body dims, the
 * Header & Footer ribbon tab appears, and Escape or a double-click on the body closes it.
 */
export function attachPrintStoryEditing(host: PrintStoryHost): () => void {
	const controller = new AbortController();
	let editing: HTMLElement | undefined;
	const finish = () => {
		editing = undefined;
		host.pages.classList.remove(EDITING);
		host.holdLayout(false);
	};
	host.pages.addEventListener(
		'dblclick',
		(event) => {
			if (!host.printActive() || !(event.target instanceof Element)) return;
			if (editing) {
				if (!event.target.closest('.dve-header-footer-editing')) {
					host.parts.closeOpenStory();
					event.stopPropagation();
					event.preventDefault();
				}
				return;
			}
			const sheet = event.target.closest<HTMLElement>('.dve-print-page');
			if (!sheet?.dataset.sectionIndex) return;
			const area = areaAt(sheet, event.clientY);
			if (!area) return;
			const sectionIndex = Number(sheet.dataset.sectionIndex);
			const slot = (sheet.dataset.headerFooterSlot ?? 'default') as HeaderFooterSlotName;
			const model = host.model();
			const stored = effectiveHeaderFooter(model, sectionIndex, area.kind, slot);
			const blocks: Block[] = stored?.blocks ?? [
				{ type: 'paragraph', id: newHeaderFooterId(), runs: [] },
			];
			const kind = area.kind === 'headers' ? 'header' : 'footer';
			let container = area.element;
			if (!container) {
				// An empty story has no element yet; give it the place a header or footer would take.
				container = renderHeaderFooter(
					{ blocks },
					{ page: '1', numPages: '1', sectionPages: '1', now: new Date() },
					kind,
				);
				container.style.left = `${sheet.dataset.marginLeft}px`;
				container.style.right = `${sheet.dataset.marginRight}px`;
				container.style[kind === 'header' ? 'top' : 'bottom'] =
					`${sheet.dataset[kind === 'header' ? 'headerDistance' : 'footerDistance']}px`;
				sheet.append(container);
			}
			host.holdLayout(true);
			const target = container;
			const opened = host.parts.editOnPage({
				container: target,
				kind: area.kind,
				slot,
				sectionIndex,
				blocks,
				closed: finish,
			});
			if (!opened) {
				host.holdLayout(false);
				if (!area.element) target.remove();
				return;
			}
			editing = target;
			host.pages.classList.add(EDITING);
			event.preventDefault();
			event.stopPropagation();
		},
		{ capture: true, signal: controller.signal },
	);
	// While a story is open the body cannot take focus or a caret; the editor keeps it.
	host.pages.addEventListener(
		'pointerdown',
		(event) => {
			if (!editing || !(event.target instanceof Element)) return;
			if (event.target.closest('.dve-header-footer-editing')) return;
			event.preventDefault();
			event.stopImmediatePropagation();
		},
		{ capture: true, signal: controller.signal },
	);
	return () => {
		controller.abort();
		if (editing) host.parts.closeOpenStory();
	};
}
