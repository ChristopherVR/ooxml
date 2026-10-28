import { createSlidesPane, createSlidesPaneBulkSlideOps } from 'pptx-vanilla-viewer';
import type { LoadedPresentation, SlidesPane, Store, ViewerState } from 'pptx-vanilla-viewer';
import { duplicateSlide } from 'pptx-viewer-core';
import type { PptxSlide } from 'pptx-viewer-core';

import { t } from './demo-i18n';

export interface HostOwnedSlidesPaneOptions {
	store: Store<ViewerState>;
	getLoaded(): LoadedPresentation | undefined;
	/** Commit an open inline text edit first; false keeps the editor open (invalid edit). */
	commitInline(): boolean;
}

const newSlideId = (): string => `slide-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

/** Keep `slideNumber` equal to the 1-based position, as the full viewer does. */
const renumber = (slides: PptxSlide[]): PptxSlide[] =>
	slides.map((slide, index) =>
		slide.slideNumber === index + 1 ? slide : { ...slide, slideNumber: index + 1 },
	);

/**
 * The custom shell's slide rail, driven from the shell's own store. Slide
 * commands rewrite `store.slides`, the same path the canvas edits take, so the
 * collaboration controller publishes them to peers.
 */
export function mountHostOwnedSlidesPane(
	container: HTMLElement,
	options: HostOwnedSlidesPaneOptions,
): SlidesPane {
	const { store } = options;
	/** Run a slide edit only while editing is allowed and no inline edit is left open. */
	const edit =
		<A extends unknown[]>(run: (...args: A) => void) =>
		(...args: A): void => {
			if (store.get().editable && options.commitInline()) {
				run(...args);
			}
		};
	const setSlides = (slides: PptxSlide[], currentSlide: number): void =>
		store.set({
			slides: renumber(slides),
			currentSlide,
			selectedElementId: null,
			selectedElementIds: [],
		});
	const addSlideAfter = (index: number): void => {
		const { slides } = store.get();
		const insertAt = Math.min(index + 1, slides.length);
		const slide: PptxSlide = { id: newSlideId(), rId: '', slideNumber: insertAt + 1, elements: [] };
		setSlides([...slides.slice(0, insertAt), slide, ...slides.slice(insertAt)], insertAt);
	};
	const bulk = createSlidesPaneBulkSlideOps({
		duplicateSlide: (index) => {
			const { slides } = store.get();
			const source = slides[index];
			if (!source) {
				return;
			}
			const copy = { ...duplicateSlide(source, slides.length + 1), id: newSlideId() };
			setSlides([...slides.slice(0, index + 1), copy, ...slides.slice(index + 1)], index + 1);
		},
		deleteSlide: (index) => {
			const { slides, currentSlide } = store.get();
			if (slides.length <= 1 || !slides[index]) {
				return;
			}
			const next = slides.filter((_, i) => i !== index);
			setSlides(
				next,
				Math.min(currentSlide > index ? currentSlide - 1 : currentSlide, next.length - 1),
			);
		},
		toggleSlideHidden: (index) => {
			const { slides, currentSlide } = store.get();
			if (slides[index]) {
				setSlides(
					slides.map((slide, i) => (i === index ? { ...slide, hidden: !slide.hidden } : slide)),
					currentSlide,
				);
			}
		},
	});
	const state = store.get();
	const pane = createSlidesPane(container, {
		slides: state.slides,
		activeSlideIndex: state.currentSlide,
		canvasSize: state.canvasSize,
		canEdit: state.editable,
		t,
		onSelectSlide: (index) => {
			if (index !== store.get().currentSlide && options.commitInline()) {
				store.set({ currentSlide: index, selectedElementId: null, selectedElementIds: [] });
			}
		},
		onAddSlide: edit(() => addSlideAfter(store.get().currentSlide)),
		onAddSlideAfter: edit(addSlideAfter),
		onDuplicateSlides: edit(bulk.duplicateSlides),
		onDeleteSlides: edit(bulk.deleteSlides),
		onHideSlides: edit(bulk.toggleHideSlides),
	});
	let loaded: LoadedPresentation | undefined;
	const unsubscribe = store.subscribe((next) => {
		const current = options.getLoaded();
		const deck =
			current && current !== loaded
				? {
						mediaDataUrls: current.mediaDataUrls,
						headerFooter: current.headerFooter,
						customProperties: current.customProperties,
						firstSlideNumber: current.firstSlideNumber,
						colorScheme: current.colorScheme,
						fontScheme: current.fontScheme,
						formatScheme: current.formatScheme,
						tableStyleMap: current.tableStyleMap,
					}
				: {};
		loaded = current;
		pane.update({
			...deck,
			slides: next.slides,
			activeSlideIndex: next.currentSlide,
			canvasSize: next.canvasSize,
			canEdit: next.editable,
		});
	});
	return {
		get el() {
			return pane.el;
		},
		update: pane.update,
		destroy() {
			unsubscribe();
			pane.destroy();
		},
	};
}
