import {
	createSlideSorterController,
	HIDDEN_SLIDE_ATTRIBUTE,
	HIDDEN_SLIDE_LABEL_KEY,
	hiddenSlideCue,
	slideClipboardLabels,
	slideSorterView,
} from 'ooxml-ui/pptx';
import type { SlideSorterState, SlidePasteRequest } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

import type { Translator } from '../i18n';
import { createEl } from '../render';
import { openSlideSorterContextMenu } from './slide-sorter-context-menu';

export interface SlideSorterOptions {
	slides: readonly PptxSlide[];
	clipboardTemplates?: Record<string, PptxElement[]>;
	current: number;
	renderSlide?(slide: PptxSlide, scale: number): HTMLElement;
	onSelect(index: number): void;
	onReorder(from: number, to: number): void;
	onDelete(index: number): void;
	onDuplicate(index: number): void;
	onPaste?(request: SlidePasteRequest): void;
	onCopySystem?(indexes: number[]): Promise<void>;
	onPasteSystem?(insertAt: number): Promise<void>;
	onUndo?(): void;
	onRedo?(): void;
	onToggleHidden(index: number): void;
	/** Whether the host allows edits; gates the deck-writing shortcuts. */
	canEdit?: boolean;
	/**
	 * Slide ids copied in the sorter. The caller keeps the object when it
	 * re-opens the overlay after an edit, so Copy then Paste survives the
	 * re-render.
	 */
	clipboard?: { ids: string[]; state?: SlideSorterState };
}

export function openSlideSorterOverlay(
	doc: Document,
	host: HTMLElement,
	t: Translator,
	options: SlideSorterOptions,
): void {
	host.querySelector('[data-pptx-slide-sorter]')?.remove();
	const overlay = createEl(doc, 'section', 'pptxv-slide-sorter');
	overlay.dataset.pptxSlideSorter = 'true';
	overlay.setAttribute('role', 'dialog');
	const header = createEl(doc, 'header');
	const title = createEl(doc, 'h2');
	title.textContent = t('pptx.slideSorter.title');
	const count = createEl(doc, 'span');
	count.textContent = t('pptx.slideSorter.slideCount', { count: options.slides.length });
	const close = createEl(doc, 'button');
	close.type = 'button';
	close.textContent = '×';
	close.setAttribute('aria-label', t('pptx.slideSorter.close'));
	const clipboardControl = doc.createElement(
		'pptx-ui-slide-clipboard',
	) as import('ooxml-ui/pptx').PptxUiSlideClipboardElement;
	clipboardControl.labels = slideClipboardLabels(t);
	clipboardControl.copyAction = options.onCopySystem
		? () => options.onCopySystem!(controller.selectedIndexes())
		: undefined;
	clipboardControl.pasteAction =
		options.canEdit !== false && options.onPasteSystem
			? () => options.onPasteSystem!(controller.pasteInsertAt())
			: undefined;
	header.append(title, count, clipboardControl, close);
	overlay.appendChild(header);
	const grid = createEl(doc, 'div', 'pptxv-sorter-grid');

	// Slide ids copied in this sorter session (the shared Copy / Paste pair).
	const clipboard = options.clipboard ?? { ids: [] };
	let closeMenu: (() => void) | null = null;
	const controller = createSlideSorterController(
		{
			slides: () => options.slides,
			activeIndex: () => options.current,
			canEdit: () => options.canEdit !== false,
			clipboardTemplates: () => options.clipboardTemplates,
			onOperation: (operation, indexes) => {
				for (const index of indexes) {
					if (operation === 'delete') {
						options.onDelete(index);
					} else if (operation === 'duplicate') {
						options.onDuplicate(index);
					} else {
						options.onToggleHidden(index);
					}
				}
			},
			onPaste: (request) => options.onPaste?.(request),
			onMove: (from, to) => options.onReorder(from, to),
			onClose: () => dismiss(),
			onUndo: options.onUndo,
			onRedo: options.onRedo,
		},
		{ initial: clipboard.state },
	);
	const refreshSelection = (): void => {
		const { sorter } = controller.getSnapshot();
		clipboard.state = sorter;
		clipboard.ids = sorter.clipboardIds;
		const view = slideSorterView(controller.getSnapshot(), options.slides);
		if (!controller.getSnapshot().menu) {
			closeMenu?.();
			closeMenu = null;
		}
		overlay
			.querySelectorAll<HTMLElement>('[data-pptx-chrome="sorter-tile"]')
			.forEach((card, index) => {
				const selected = view.selectedIds.has(options.slides[index].id);
				card.dataset.pptxSelected = String(selected);
				card.classList.toggle('is-current', selected);
				card.classList.toggle('is-drop-target', controller.getSnapshot().dragOverIndex === index);
			});
		const slider = overlay.querySelector<HTMLInputElement>('input[type=range]');
		if (slider) {
			slider.value = String(sorter.zoom);
		}
		grid.style.gridTemplateColumns = `repeat(${view.gridColumns}, minmax(0, 1fr))`;
		grid.querySelectorAll<HTMLElement>('article > button').forEach((button) => {
			button.style.height = `${100 * view.zoomScale}px`;
		});
	};
	const unsubscribe = controller.subscribe(refreshSelection);
	options.slides.forEach((slide, index) => {
		const card = createEl(doc, 'article', 'pptxv-sorter-card');
		card.dataset.pptxChrome = 'sorter-tile';
		card.dataset.pptxTouchPress = 'menu dblclick';
		card.addEventListener('contextmenu', (event) => {
			closeMenu?.();
			closeMenu = null;
			if (!controller.openMenu(index, event.clientX, event.clientY)) {
				return;
			}
			event.preventDefault();
			closeMenu = openSlideSorterContextMenu({
				doc,
				t,
				host: overlay,
				x: event.clientX,
				y: event.clientY,
				context: slideSorterView(controller.getSnapshot(), options.slides).menuContext,
				onCommand: (command) => controller.run(command),
			});
		});
		card.draggable = options.canEdit !== false;
		card.classList.toggle('is-current', index === options.current);
		card.classList.toggle('is-hidden', Boolean(slide.hidden));
		const preview = createEl(doc, 'button');
		preview.type = 'button';
		// The number lives in its own span so the hidden-slide slash can be drawn
		// across the number alone rather than the whole preview button.
		const num = createEl(doc, 'span', 'pptxv-sorter-num');
		num.textContent = String(index + 1);
		preview.appendChild(num);
		if (options.renderSlide) {
			preview.appendChild(options.renderSlide(slide, 0.2));
		}
		preview.setAttribute('aria-label', t('pptx.compare.slideNumber', { number: index + 1 }));
		const cue = hiddenSlideCue(slide.hidden, 'sorter', index);
		if (cue.marker && cue.labelId) {
			card.setAttribute(HIDDEN_SLIDE_ATTRIBUTE, cue.marker);
			preview.setAttribute('aria-describedby', cue.labelId);
			const badge = createEl(doc, 'span', 'pptxv-sorter-hidden');
			badge.id = cue.labelId;
			badge.textContent = t(HIDDEN_SLIDE_LABEL_KEY);
			preview.appendChild(badge);
		}
		preview.addEventListener('click', (event) => controller.select(index, event));
		preview.addEventListener('dblclick', () => {
			options.onSelect(index);
			overlay.remove();
		});
		card.addEventListener('dragstart', (event) => {
			event.dataTransfer?.setData('text/plain', String(index));
			controller.dragStart(index);
		});
		card.addEventListener('dragover', (event) => {
			event.preventDefault();
			controller.dragOver(index);
		});
		card.addEventListener('dragleave', () => controller.dragLeave());
		card.addEventListener('dragend', () => controller.dragEnd());
		card.addEventListener('drop', (event) => {
			event.preventDefault();
			const data = event.dataTransfer?.getData('text/plain');
			controller.drop(index, data ? Number(data) : undefined);
		});
		card.append(preview);
		grid.appendChild(card);
	});
	overlay.appendChild(grid);

	const dismiss = (): void => {
		closeMenu?.();
		unsubscribe();
		doc.removeEventListener('keydown', onKeyDown, true);
		overlay.remove();
	};
	const onKeyDown = (event: KeyboardEvent): void => {
		if (!overlay.isConnected) {
			unsubscribe();
			doc.removeEventListener('keydown', onKeyDown, true);
			return;
		}
		controller.handleKeydown(event);
	};
	doc.addEventListener('keydown', onKeyDown, true);

	const zoom = createEl(doc, 'input');
	zoom.type = 'range';
	zoom.min = '50';
	zoom.max = '200';
	zoom.step = '10';
	zoom.value = String(controller.getSnapshot().sorter.zoom);
	zoom.setAttribute('aria-label', t('pptx.slideSorter.zoom'));
	zoom.addEventListener('input', () => controller.setZoom(Number(zoom.value)));
	overlay.appendChild(zoom);
	refreshSelection();
	close.addEventListener('click', dismiss);
	host.appendChild(overlay);
	overlay.tabIndex = -1;
	overlay.focus();
}
