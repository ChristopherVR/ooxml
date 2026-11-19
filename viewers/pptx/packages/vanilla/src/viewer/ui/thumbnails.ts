import {
	computeVirtualRange,
	createVirtualWindow,
	focusCurrentRow,
	railHasFocus,
	EDITOR_THUMBNAIL_WIDTH,
	EDITOR_SLIDE_RAIL_WIDTH,
	editorThumbnailStep,
	SLIDE_VIRTUALIZATION_THRESHOLD,
} from 'ooxml-ui/pptx';
import type { CanvasSize } from 'ooxml-ui/pptx';
import type { PptxSection, PptxSlide, PptxSlideMaster } from 'pptx-viewer-core';

import type { EditActions } from '../editor';
import type { Translator } from '../i18n';
import { disposePptxRenderStage } from '../render';
import { collectThreeViews, createEl, withReusableThreeViews } from '../render';
import type { Store, ViewerState } from '../state';
import { createThumbnailContextMenu } from './thumbnail-context-menu';
import { wireThumbnailRowEvents } from './thumbnail-events';
import { createThumbnailRailMenu } from './thumbnail-rail-menu';
import type { ThumbnailSectionActions } from './thumbnail-sections';
import { renderThumbnailSections } from './thumbnail-sections';
import {
	createThumbnailFooter,
	createThumbnailMasterRows,
	createThumbnailRow,
} from './thumbnail-views';

export type { ThumbnailSectionActions } from './thumbnail-sections';

/** The item-2 multi-select/context-menu deps; all optional so existing callers keep working. */
export interface ThumbnailRailMenuDeps {
	/** The viewer store, for the built-in Layout list; omit with `openLayout`. */
	store?: Store<ViewerState>;
	/** Layout and Add Section edit actions; omit with `openLayout` and `addSection`. */
	getEditActions?(): EditActions;
	addSlideAfter(index: number): void;
	duplicateSlides(indexes: number[]): void;
	deleteSlides(indexes: number[]): void;
	toggleHideSlides(indexes: number[]): void;
	/** Host handler for the menu's "Layout" entry (replaces the built-in list). */
	openLayout?(index: number, x: number, y: number): void;
	/** Host handler for the menu's "Add Section" entry. */
	addSection?(name: string, afterSlideIndex: number): void;
}

const THUMB_STAGE_WIDTH = EDITOR_THUMBNAIL_WIDTH;
/**
 * `aria-current` value of the active slide row; the same token React, Vue,
 * Svelte and Angular put on their rail rows.
 */
const CURRENT_ROW = 'true';

export interface ThumbnailRail {
	el: HTMLElement;
	/** Rebuild the rail for a new slide list (uses `renderStage` per slide). */
	render(
		slides: PptxSlide[],
		canvasSize: CanvasSize,
		renderStage: (slide: PptxSlide, scale: number) => HTMLElement,
		sections?: readonly PptxSection[],
		sectionActions?: ThumbnailSectionActions,
	): void;
	/** Highlight the active slide and scroll it into view. */
	setActive(index: number): void;
	/** Show or hide the rail. */
	setVisible(visible: boolean): void;
	/** Show or hide the pinned Add Slide footer (editing enabled only). */
	setAddSlideVisible(visible: boolean): void;
	renderMasters(
		masters: readonly PptxSlideMaster[],
		canvasSize: CanvasSize,
		renderStage: (slide: PptxSlide, scale: number) => HTMLElement,
		onSelect: (masterIndex: number, layoutIndex: number | null) => void,
		active: { masterIndex: number; layoutIndex: number | null },
	): void;
	/** Close an open thumbnail menu and release the rendered thumbnail stages. */
	destroy(): void;
}

/**
 * The thumbnail sidebar: a scaled-down live render of every slide; clicking a
 * thumbnail navigates to it. Every row is rebuilt on `render`; scrolling and
 * selection build only the rows entering the window.
 */
export function createThumbnailRail(
	doc: Document,
	t: Translator,
	onSelect: (index: number) => void,
	onAddSlide?: () => void,
	menuDeps?: ThumbnailRailMenuDeps,
): ThumbnailRail {
	const el = createEl(doc, 'aside', 'pptxv-thumbs');
	el.setAttribute('role', 'navigation');
	el.dataset.pptxChrome = 'slides';
	el.style.width = `${EDITOR_SLIDE_RAIL_WIDTH}px`;
	el.setAttribute('aria-label', t('pptx.sections.slides'));

	// ── Ctrl/Shift multi-select + thumbnail right-click menu ────────────────
	const railMenu = createThumbnailRailMenu();
	const contextMenu = menuDeps
		? createThumbnailContextMenu({
				doc,
				store: menuDeps.store,
				getTranslator: () => t,
				getEditActions: menuDeps.getEditActions,
				addSlideAfter: menuDeps.addSlideAfter,
				duplicateSlides: menuDeps.duplicateSlides,
				deleteSlides: menuDeps.deleteSlides,
				toggleHideSlides: menuDeps.toggleHideSlides,
				openLayout: menuDeps.openLayout,
				addSection: menuDeps.addSection,
				host: el,
			})
		: null;
	// PowerPoint's Enter on a focused thumbnail inserts a new slide after it.
	el.addEventListener('keydown', (event) => {
		if (event.key !== 'Enter' || !menuDeps) {
			return;
		}
		const target = event.target;
		if (target instanceof HTMLElement && target.dataset.slideIndex !== undefined) {
			event.preventDefault();
			menuDeps.addSlideAfter(activeIndex);
		}
	});
	// Scrollable slide list; the Add Slide footer stays pinned below it
	// (mirrors React's SlidesPaneSidebar bottom button).
	const list = createEl(doc, 'div', 'pptxv-thumbs-list');
	list.dataset.pptxChrome = 'slide-list';
	el.appendChild(list);
	let footer: HTMLElement | null = null;
	let addSlideVisible = false;
	let masterMode = false;
	const applyFooterVisibility = (): void => {
		if (footer) {
			footer.hidden = !addSlideVisible || masterMode;
		}
	};
	if (onAddSlide) {
		footer = createThumbnailFooter(doc, t, onAddSlide);
		el.appendChild(footer);
	}
	let buttons = new Map<number, HTMLButtonElement>();
	/** Window rows by slide id; a row is reused while its slide and index hold. */
	let rows = new Map<
		string,
		{ slide: PptxSlide; index: number; el: HTMLElement; button: HTMLButtonElement }
	>();
	let activeIndex = 0;
	let sourceSlides: PptxSlide[] = [];
	let sourceCanvasSize: CanvasSize = { width: 1, height: 1 };
	let sourceRenderStage: ((slide: PptxSlide, scale: number) => HTMLElement) | null = null;
	let sourceSections: readonly PptxSection[] = [];
	let itemHeight = 1;
	let virtualized = false;
	const space = createEl(doc, 'div', 'pptxv-thumbs-space', { position: 'relative' });
	const window = createEl(doc, 'div', 'pptxv-thumbs-window', {
		display: 'flex',
		flexDirection: 'column',
	});
	window.dataset.pptxChrome = 'slide-window';
	space.appendChild(window);
	const virtualWindow = createVirtualWindow({ totalItems: 0, itemHeight: 1 }, () => renderWindow());
	/**
	 * Rows entering a virtualized window while scrolling get their thumbnail
	 * once scrolling pauses; an explicit render paints every row at once.
	 */
	let deferEntering = true;
	const deferredStage = (slide: PptxSlide, scale: number): HTMLElement => {
		const pending = createEl(doc, 'span', 'pptxv-thumb-pending');
		virtualWindow.defer(slide.id, () => {
			if (pending.isConnected && sourceRenderStage) {
				pending.replaceWith(sourceRenderStage(slide, scale));
			}
		});
		return pending;
	};

	const buildButton = (slide: PptxSlide, index: number, scale: number): HTMLButtonElement => {
		const btn = createThumbnailRow(
			doc,
			t,
			slide,
			index,
			sourceCanvasSize,
			virtualized && deferEntering ? deferredStage : sourceRenderStage!,
			scale,
		);
		wireThumbnailRowEvents(btn, slide, index, {
			selection: railMenu,
			menu: contextMenu,
			getSlides: () => sourceSlides,
			getActive: () => activeIndex,
			onSelect,
		});
		buttons.set(index, btn);
		return btn;
	};
	/** A row of the window: in a virtualized rail it says where it sits in the deck. */
	const rowOf = (button: HTMLButtonElement, index: number): HTMLElement => {
		if (!virtualized) {
			return button;
		}
		const item = createEl(doc, 'div', 'pptxv-thumb-item');
		item.setAttribute('role', 'listitem');
		item.setAttribute('aria-setsize', String(sourceSlides.length));
		item.setAttribute('aria-posinset', String(index + 1));
		item.appendChild(button);
		return item;
	};

	const markActive = (): void => {
		buttons.forEach((btn, buttonIndex) => {
			const active = buttonIndex === activeIndex;
			btn.classList.toggle('is-active', active);
			if (active) {
				btn.setAttribute('aria-current', CURRENT_ROW);
			} else {
				btn.removeAttribute('aria-current');
			}
		});
	};

	/** Drop every row, so the next window builds each one again. */
	const clearRows = (): void => {
		disposePptxRenderStage(list);
		for (const id of rows.keys()) {
			virtualWindow.cancelDeferred(id);
		}
		rows = new Map();
		buttons = new Map();
	};

	/**
	 * Show the rows of the current window: rows still in it stay in place,
	 * rows that left are released, and only the entering rows are built.
	 */
	const renderWindow = (): void => {
		if (!sourceRenderStage || sourceSections.length > 0) {
			return;
		}
		const scale = THUMB_STAGE_WIDTH / Math.max(sourceCanvasSize.width, 1);
		const range = virtualized
			? virtualWindow.range
			: computeVirtualRange(
					sourceSlides.length,
					itemHeight,
					0,
					sourceSlides.length * itemHeight,
					0,
				);
		const wanted = new Map<string, number>();
		for (let index = range.startIndex; index <= range.endIndex; index += 1) {
			const slide = sourceSlides[index];
			if (slide) {
				wanted.set(slide.id, index);
			}
		}
		const leaving: HTMLElement[] = [];
		for (const [id, row] of rows) {
			if (wanted.get(id) !== row.index || sourceSlides[row.index] !== row.slide) {
				leaving.push(row.el);
				rows.delete(id);
				virtualWindow.cancelDeferred(id);
			}
		}
		// Only views of leaving rows may move into entering rows; a kept row's
		// view must stay where it is.
		const reusableViews: ReturnType<typeof collectThreeViews> = new Map();
		for (const row of leaving) {
			disposePptxRenderStage(row);
			for (const [spec, view] of collectThreeViews(row)) {
				if (!reusableViews.has(spec)) {
					reusableViews.set(spec, view);
				}
			}
		}
		buttons = new Map();
		const ordered: HTMLElement[] = [];
		const build = (): void => {
			for (const [id, index] of wanted) {
				let row = rows.get(id);
				if (!row) {
					const slide = sourceSlides[index]!;
					const button = buildButton(slide, index, scale);
					row = { slide, index, el: rowOf(button, index), button };
					rows.set(id, row);
				}
				buttons.set(index, row.button);
				ordered.push(row.el);
			}
		};
		// With no leaving rows, keep the caller's pool (a full render's views).
		if (reusableViews.size > 0) {
			withReusableThreeViews(reusableViews, build);
		} else {
			build();
		}
		for (const row of leaving) {
			row.remove();
		}
		// Kept rows are already in order; each entering row goes before the
		// first kept row after it, so no kept row moves.
		let cursor = window.firstChild;
		for (const row of ordered) {
			if (row === cursor) {
				cursor = cursor.nextSibling;
			} else {
				window.insertBefore(row, cursor);
			}
		}
		window.style.top = virtualized ? `${range.offsetY}px` : '';
		space.style.height = virtualized ? `${range.totalHeight}px` : 'auto';
		markActive();
	};

	return {
		el,
		render(slides, canvasSize, renderStage, sections, sectionActions) {
			masterMode = false;
			applyFooterVisibility();
			sourceSlides = slides;
			sourceCanvasSize = canvasSize;
			sourceRenderStage = renderStage;
			sourceSections = sections ?? [];
			itemHeight = editorThumbnailStep(canvasSize.width, canvasSize.height);
			virtualized = !sections?.length && slides.length >= SLIDE_VIRTUALIZATION_THRESHOLD;
			if (virtualized) {
				window.setAttribute('role', 'list');
			} else {
				window.removeAttribute('role');
			}
			// Class toggle (not an inline display) so the presenting-mode and
			// mobile-layout `display: none` stylesheet rules can still hide the
			// rail; an inline style would override them and leak thumbnail text
			// into presentation mode.
			el.classList.toggle('pptxv-thumbs-virtualized', virtualized);
			// An explicit render repaints every row: fonts, fields and table
			// styles reach the thumbnails without changing the slide objects.
			const views = collectThreeViews(list);
			const scale = THUMB_STAGE_WIDTH / Math.max(canvasSize.width, 1);
			deferEntering = false;
			withReusableThreeViews(views, () => {
				clearRows();
				if (sourceSections.length > 0) {
					virtualWindow.attach(null);
					list.replaceChildren(
						...renderThumbnailSections({
							doc,
							t,
							sections: sourceSections,
							slides,
							actions: sectionActions,
							buildSlide: (slide, index) => buildButton(slide, index, scale),
						}),
					);
					markActive();
					return;
				}
				window.replaceChildren();
				window.style.position = virtualized ? 'absolute' : '';
				window.style.insetInline = virtualized ? '0' : '';
				if (virtualized) {
					space.dataset.virtualized = 'true';
				} else {
					delete space.dataset.virtualized;
				}
				if (list.firstChild !== space || list.childNodes.length !== 1) {
					list.replaceChildren(space);
				}
				virtualWindow.attach(virtualized ? list : null);
				virtualWindow.setShape({ totalItems: slides.length, itemHeight });
				renderWindow();
			});
			deferEntering = true;
			this.setActive(activeIndex);
			// Rebuilding the rows drops the focused one: focus moves to the current row rather
			// than falling to the page.
			if (railHasFocus(list)) {
				focusCurrentRow(list);
			}
		},
		setActive(index) {
			const changed = index !== activeIndex;
			activeIndex = index;
			// Checked first: the reveal can remove the row that holds focus.
			const keepFocus = changed && railHasFocus(list);
			if (virtualized) {
				// Builds rows only when the reveal moved the window.
				virtualWindow.reveal(index);
			}
			if (changed || !buttons.get(index)?.hasAttribute('aria-current')) {
				markActive();
			}
			if (keepFocus) {
				focusCurrentRow(list);
			}
			const active = buttons.get(index);
			if (!virtualized && active && typeof active.scrollIntoView === 'function') {
				active.scrollIntoView({ block: 'nearest' });
			}
		},
		setVisible(visible) {
			el.hidden = !visible;
		},
		setAddSlideVisible(visible) {
			addSlideVisible = visible;
			applyFooterVisibility();
		},
		destroy() {
			contextMenu?.destroy();
			virtualWindow.destroy();
			clearRows();
			list.replaceChildren();
			sourceRenderStage = null;
		},
		renderMasters(masters, canvasSize, renderStage, select, active) {
			masterMode = true;
			applyFooterVisibility();
			clearRows();
			list.replaceChildren();
			virtualWindow.attach(null);
			virtualized = false;
			el.classList.remove('pptxv-thumbs-virtualized');
			list.append(
				...createThumbnailMasterRows(doc, t, masters, canvasSize, renderStage, select, active),
			);
		},
	};
}
