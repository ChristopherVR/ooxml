import { createSlideRailController, slideRailMenuEntries } from 'ooxml-ui/pptx';
import type { SlidePaneContextMenuCommandId } from 'ooxml-ui/pptx';
/**
 * useSlidePaneRailMenu: Ctrl/Shift multi-select and the thumbnail right-click
 * menu for `SlidesPaneSidebar.vue`.
 *
 * Split out of the SFC (already past the repo's 300-line guideline once this
 * landed) so the component stays template plus thin wiring. The selection,
 * the menu and its command routing are the shared `createSlideRailController`;
 * this module only reads it through `useController`.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import { computed } from 'vue';
import type { ComputedRef } from 'vue';
import { useI18n } from 'vue-i18n';

import type { ContextMenuItem } from '../components/ContextMenu.vue';
import { useController } from './useController';

export interface SlidePaneRailMenuEmit {
	select: (index: number) => void;
	'add-slide-after': (index: number) => void;
	duplicate: (indexes: number[]) => void;
	delete: (indexes: number[]) => void;
	'toggle-hidden': (indexes: number[]) => void;
	layout: (payload: { index: number; x: number; y: number }) => void;
	'add-section': (index: number) => void;
}

export interface SlidePaneContextMenuState {
	open: boolean;
	x: number;
	y: number;
	index: number;
	selectedIndexes: number[];
}

export interface UseSlidePaneRailMenuResult {
	selectedIds: ComputedRef<readonly string[]>;
	isSelected: (slideId: string) => boolean;
	onSlideClick: (e: MouseEvent, index: number) => void;
	onPaneKeydown: (e: KeyboardEvent) => void;
	menu: ComputedRef<SlidePaneContextMenuState>;
	closeMenu: () => void;
	menuItems: ComputedRef<ContextMenuItem[]>;
	onContextMenu: (e: MouseEvent, index: number) => void;
	onMenuSelect: (id: string) => void;
}

export function useSlidePaneRailMenu(
	slides: () => readonly PptxSlide[],
	activeIndex: () => number,
	canEdit: () => boolean,
	emit: <K extends keyof SlidePaneRailMenuEmit>(
		event: K,
		...args: Parameters<SlidePaneRailMenuEmit[K]>
	) => void,
): UseSlidePaneRailMenuResult {
	const { t } = useI18n();

	// The selection (by id: an index shifts under insert, delete and move), the menu and its
	// command routing are the shared slide rail controller.
	const rail = createSlideRailController({ canEdit });
	const state = useController(rail);
	const selectedIds = computed(() => state.value.selectedIds);
	function isSelected(slideId: string): boolean {
		return state.value.selectedIds.includes(slideId);
	}
	const orderedIds = (): string[] => slides().map((s) => s.id);
	function onSlideClick(e: MouseEvent, index: number): void {
		const slide = slides()[index];
		if (slide) {
			rail.click(e, slide.id, orderedIds());
		}
		emit('select', index);
	}
	/** PowerPoint's Enter on a focused thumbnail inserts a new slide after it. */
	function onPaneKeydown(e: KeyboardEvent): void {
		if (e.key === 'Enter' && canEdit()) {
			e.preventDefault();
			emit('add-slide-after', activeIndex());
		}
	}

	// ── Slide context menu ──
	const menu = computed<SlidePaneContextMenuState>(() => {
		const open = state.value.menu;
		return open
			? { open: true, ...open, selectedIndexes: [...open.selectedIndexes] }
			: { open: false, x: 0, y: 0, index: -1, selectedIndexes: [] };
	});
	const menuItems = computed<ContextMenuItem[]>(() => {
		const count = state.value.menu?.selectedIndexes.length ?? 0;
		return slideRailMenuEntries(state.value, slides()).flatMap((entry, index) => {
			const item: ContextMenuItem = {
				id: entry.id,
				label: entry.countLabelKey ? t(entry.labelKey, { count }) : t(entry.labelKey),
				disabled: entry.disabled,
				danger: entry.id === 'delete',
			};
			return entry.separatorBefore
				? [{ id: `sep-${index}`, label: '', separator: true }, item]
				: [item];
		});
	});
	function onContextMenu(e: MouseEvent, index: number): void {
		if (rail.openMenu(e.clientX, e.clientY, index, orderedIds())) {
			e.preventDefault();
		}
	}
	function onMenuSelect(id: string): void {
		rail.run(id as SlidePaneContextMenuCommandId, {
			addSlideAfter: (index) => emit('add-slide-after', index),
			duplicateSlides: (indexes) => emit('duplicate', indexes),
			deleteSlides: (indexes) => emit('delete', indexes),
			openLayoutForSlide: (index, x, y) => emit('layout', { index, x, y }),
			toggleHideSlides: (indexes) => emit('toggle-hidden', indexes),
			addSectionAt: (index) => emit('add-section', index),
		});
	}

	return {
		selectedIds,
		isSelected,
		onSlideClick,
		onPaneKeydown,
		menu,
		closeMenu: rail.closeMenu,
		menuItems,
		onContextMenu,
		onMenuSelect,
	};
}
