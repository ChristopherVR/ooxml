import type { SlidePaneContextMenuCommandId, SlideRailCommands } from 'ooxml-ui/pptx';
import { createSlideRailController, slideRailMenuEntries } from 'ooxml-ui/pptx';
/**
 * thumbnail-rail-menu.svelte.ts: Ctrl/Shift multi-select and the thumbnail
 * right-click menu for `ThumbnailRail.svelte`.
 *
 * Split out of the SFC (already past the repo's 300-line guideline) so the
 * component stays template plus thin wiring. The selection, the menu and its
 * command routing are the shared `createSlideRailController`; this module only
 * reads it through `controllerState`.
 */
import type { PptxSlide } from 'pptx-viewer-core';

import { controllerState } from '../state/controller-state';

/** Position and target of an open thumbnail context menu. */
export interface ThumbnailContextMenuState {
	x: number;
	y: number;
	index: number;
	selectedIndexes: number[];
}

export interface ThumbnailRailMenuActions {
	addSlideAfter?: (index: number) => void;
	duplicateSlides?: (indexes: number[]) => void;
	deleteSlides?: (indexes: number[]) => void;
	openLayoutForSlide?: (index: number, x: number, y: number) => void;
	toggleHideSlides?: (indexes: number[]) => void;
	addSectionAt?: (index: number) => void;
}

export class ThumbnailRailMenu {
	readonly controller;
	readonly #state;

	/** `canEdit` gates the menu (a read-only deck keeps the browser's own). */
	constructor(canEdit: () => boolean = () => true) {
		this.controller = createSlideRailController({ canEdit });
		this.#state = controllerState(this.controller);
	}

	get selectedIds(): string[] {
		return [...this.#state.current.selectedIds];
	}

	get contextMenu(): ThumbnailContextMenuState | null {
		const menu = this.#state.current.menu;
		return menu ? { ...menu, selectedIndexes: [...menu.selectedIndexes] } : null;
	}

	isSelected(slideId: string): boolean {
		return this.#state.current.selectedIds.includes(slideId);
	}

	/** One click on a thumbnail: the shared Ctrl/Shift rule. */
	onClick(
		event: { ctrlKey: boolean; metaKey: boolean; shiftKey: boolean },
		slideId: string,
		orderedIds: readonly string[],
	): void {
		this.controller.click(event, slideId, orderedIds);
	}

	/** A right-click: the whole selection when inside it, else that slide; false when read-only. */
	openContextMenu(x: number, y: number, index: number, orderedIds: readonly string[]): boolean {
		return this.controller.openMenu(x, y, index, orderedIds);
	}

	closeContextMenu(): void {
		this.controller.closeMenu();
	}

	menuEntries(slides: readonly PptxSlide[]) {
		return slideRailMenuEntries(this.#state.current, slides);
	}

	/** Run a chosen command id against `actions`, then close the menu. */
	run(id: SlidePaneContextMenuCommandId, actions: ThumbnailRailMenuActions): void {
		this.controller.run(id, actions satisfies SlideRailCommands);
	}
}
