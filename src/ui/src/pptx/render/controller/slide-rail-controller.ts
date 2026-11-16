/**
 * The slide rail (the thumbnails pane) as one state owner for every binding: the Ctrl/Shift
 * multi-selection of thumbnails with its anchor, and the right-click menu with the slides it acts
 * on.
 *
 * The click rule (`resolveSlidePaneClick`) and the menu's entries
 * (`buildSlidePaneContextMenuEntries`) were already shared; each binding still kept its own copy of
 * the state around them (Angular's `SlidePaneRailSelection`, Svelte's `ThumbnailRailMenu`, Vue's
 * `useSlidePaneRailMenu`, React's pane callbacks, vanilla's rail menu). Now: the selection is kept
 * by slide id (an index shifts under insert, delete and move), a right-click inside the selection
 * acts on the whole selection and outside it on that slide alone, the menu is refused on a
 * read-only deck, and a chosen command runs on the binding's own slide commands.
 *
 * @module render/controller/slide-rail-controller
 */
import type { PptxSlide } from 'ooxml-core/pptx';

import { buildSlidePaneContextMenuEntries } from '../slide-pane-context-menu';
import type {
	SlidePaneContextMenuCommandId,
	SlidePaneContextMenuEntry,
} from '../slide-pane-context-menu';
import { resolveSlidePaneClick } from '../slide-pane-selection';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

/** Where the menu is open, over which thumbnail, and every slide it acts on. */
export interface SlideRailMenuState {
	readonly x: number;
	readonly y: number;
	readonly index: number;
	/** Every slide the menu's bulk commands act on, the right-clicked one included. */
	readonly selectedIndexes: readonly number[];
}

export interface SlideRailState {
	/** Selected thumbnails, by slide id. */
	readonly selectedIds: readonly string[];
	/** Where a Shift+click range starts. */
	readonly anchorId: string | null;
	readonly menu: SlideRailMenuState | null;
}

/** The binding's slide commands the menu runs (each optional: a missing one does nothing). */
export interface SlideRailCommands {
	addSlideAfter?(index: number): void;
	duplicateSlides?(indexes: number[]): void;
	deleteSlides?(indexes: number[]): void;
	openLayoutForSlide?(index: number, x: number, y: number): void;
	toggleHideSlides?(indexes: number[]): void;
	addSectionAt?(index: number): void;
}

export interface SlideRailControllerDeps {
	/** Whether the host allows edits: the menu is refused otherwise. */
	canEdit(): boolean;
}

/** Click modifiers of a thumbnail click. */
export interface SlideRailClick {
	readonly ctrlKey: boolean;
	readonly metaKey: boolean;
	readonly shiftKey: boolean;
}

export interface SlideRailController extends PptxController<SlideRailState> {
	/** A click on `slideId`: Ctrl/Cmd toggles, Shift extends from the anchor, plain selects. */
	click(event: SlideRailClick, slideId: string, orderedIds: readonly string[]): void;
	/** Replace the selection (the editor moved to another slide, a host select). */
	setSelection(ids: readonly string[]): void;
	/** A right-click on thumbnail `index`; returns false on a read-only deck. */
	openMenu(x: number, y: number, index: number, orderedIds: readonly string[]): boolean;
	closeMenu(): void;
	/** Run a menu command on `commands`, then close the menu. */
	run(id: SlidePaneContextMenuCommandId, commands: SlideRailCommands): void;
}

/** The menu's entries for the slides it acts on. */
export function slideRailMenuEntries(
	state: SlideRailState,
	slides: readonly PptxSlide[],
): SlidePaneContextMenuEntry[] {
	const selected = (state.menu?.selectedIndexes ?? [])
		.map((index) => slides[index])
		.filter((slide): slide is PptxSlide => Boolean(slide));
	return buildSlidePaneContextMenuEntries({
		selectedCount: selected.length,
		hasHiddenInSelection: selected.some((slide) => slide.hidden),
		hasVisibleInSelection: selected.some((slide) => !slide.hidden),
		wouldDeleteAllSlides: selected.length >= slides.length,
	});
}

export function createSlideRailController(deps: SlideRailControllerDeps): SlideRailController {
	const store = createControllerStore<SlideRailState>({
		selectedIds: [],
		anchorId: null,
		menu: null,
	});
	const closeMenu = (): void =>
		store.update((current) => (current.menu ? { ...current, menu: null } : current));
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		click(event, slideId, orderedIds) {
			const current = store.getSnapshot();
			const result = resolveSlidePaneClick({
				clickedId: slideId,
				orderedIds,
				selectedIds: [...current.selectedIds],
				anchorId: current.anchorId,
				ctrlKey: event.ctrlKey,
				metaKey: event.metaKey,
				shiftKey: event.shiftKey,
			});
			store.set({ ...current, selectedIds: result.selectedIds, anchorId: result.anchorId });
		},
		setSelection(ids) {
			store.update((current) => ({ ...current, selectedIds: [...ids], anchorId: ids[0] ?? null }));
		},
		openMenu(x, y, index, orderedIds) {
			const slideId = orderedIds[index];
			if (!deps.canEdit() || slideId === undefined) {
				return false;
			}
			const current = store.getSnapshot();
			const ids = current.selectedIds.includes(slideId) ? current.selectedIds : [slideId];
			const selectedIndexes = ids.map((id) => orderedIds.indexOf(id)).filter((i) => i !== -1);
			store.set({ ...current, menu: { x, y, index, selectedIndexes } });
			return true;
		},
		closeMenu,
		run(id, commands) {
			const menu = store.getSnapshot().menu;
			closeMenu();
			if (!menu) {
				return;
			}
			const indexes = [...menu.selectedIndexes];
			switch (id) {
				case 'new-slide':
					commands.addSlideAfter?.(menu.index);
					break;
				case 'duplicate':
					commands.duplicateSlides?.(indexes);
					break;
				case 'delete':
					commands.deleteSlides?.(indexes);
					break;
				case 'layout':
					commands.openLayoutForSlide?.(menu.index, menu.x, menu.y);
					break;
				case 'hide':
					commands.toggleHideSlides?.(indexes);
					break;
				case 'add-section':
					commands.addSectionAt?.(menu.index);
					break;
			}
		},
	};
}
