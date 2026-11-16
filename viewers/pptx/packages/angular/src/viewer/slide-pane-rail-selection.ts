/**
 * slide-pane-rail-selection.ts: the slides pane rail's Ctrl/Shift multi-select and right-click
 * menu, as signals over the shared `createSlideRailController` (`ooxml-ui/pptx`).
 *
 * A plain class (no DI, no component) so it is unit-testable directly: this package has no
 * TestBed (see `vitest.config.ts`), and Angular `signal()`s work outside an injection context.
 *
 * @module angular-viewer/slide-pane-rail-selection
 */
import { computed, signal } from '@angular/core';
import type { Signal } from '@angular/core';
import { createSlideRailController } from 'ooxml-ui/pptx';
import type {
	SlidePaneContextMenuCommandId,
	SlideRailCommands,
	SlideRailState,
} from 'ooxml-ui/pptx';

/** Position and target of an open thumbnail context menu. */
export interface SlidePaneContextMenuState {
	x: number;
	y: number;
	index: number;
	/** Every selected slide index, right-clicked one included, for bulk commands. */
	selectedIndexes: number[];
}

export class SlidePaneRailSelection {
	readonly controller;
	readonly #state = signal<SlideRailState | null>(null);
	readonly selectedIds: Signal<string[]> = computed(() => [...this.#snapshot().selectedIds]);
	readonly contextMenu: Signal<SlidePaneContextMenuState | null> = computed(() => {
		const menu = this.#snapshot().menu;
		return menu ? { ...menu, selectedIndexes: [...menu.selectedIndexes] } : null;
	});

	/** `canEdit` gates the menu (a read-only deck keeps the browser's own). */
	constructor(canEdit: () => boolean = () => true) {
		this.controller = createSlideRailController({ canEdit });
		this.#state.set(this.controller.getSnapshot());
		this.controller.subscribe(() => this.#state.set(this.controller.getSnapshot()));
	}

	#snapshot(): SlideRailState {
		return this.#state() ?? this.controller.getSnapshot();
	}

	isSelected(slideId: string): boolean {
		return this.selectedIds().includes(slideId);
	}

	/** Resolve one click on a slide thumbnail into the next multi-selection. */
	onClick(
		event: { ctrlKey: boolean; metaKey: boolean; shiftKey: boolean },
		slideId: string,
		orderedIds: readonly string[],
	): void {
		this.controller.click(event, slideId, orderedIds);
	}

	/**
	 * Open the context menu for a right-click on `index`. A slide already part of the selection
	 * acts with the whole selection; a slide right-clicked outside it acts on just that one.
	 * Returns false on a read-only deck.
	 */
	openContextMenu(x: number, y: number, index: number, orderedIds: readonly string[]): boolean {
		return this.controller.openMenu(x, y, index, orderedIds);
	}

	closeContextMenu(): void {
		this.controller.closeMenu();
	}

	/** Run a menu command on the pane's slide commands, then close the menu. */
	run(id: SlidePaneContextMenuCommandId, commands: SlideRailCommands): void {
		this.controller.run(id, commands);
	}
}
