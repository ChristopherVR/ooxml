/**
 * thumbnail-rail-menu.ts: Ctrl/Shift multi-select and right-click menu
 * position/target state for the slides pane rail (`createThumbnailRail`).
 *
 * The selection and the menu's target are the shared `createSlideRailController`
 * (`ooxml-ui/pptx`); this factory keeps the rail's call shape over it. Kept apart
 * from `thumbnails.ts` so that file stays DOM assembly, not selection logic.
 *
 * @module ui/thumbnail-rail-menu
 */
import { createSlideRailController } from 'ooxml-ui/pptx';

/** Position and target of an open thumbnail context menu. */
export interface ThumbnailContextMenuState {
	x: number;
	y: number;
	index: number;
	/** Every selected slide index, right-clicked one included, for bulk commands. */
	selectedIndexes: number[];
}

export interface ThumbnailRailMenu {
	isSelected(slideId: string): boolean;
	onClick(
		event: { ctrlKey: boolean; metaKey: boolean; shiftKey: boolean },
		slideId: string,
		orderedIds: readonly string[],
	): void;
	/**
	 * Resolve a right-click on `index` into the menu's target: a slide already
	 * part of the selection acts with the whole selection; a slide right-
	 * clicked outside it acts on just that one.
	 */
	openContextMenu(
		x: number,
		y: number,
		index: number,
		orderedIds: readonly string[],
	): ThumbnailContextMenuState | null;
}

export function createThumbnailRailMenu(): ThumbnailRailMenu {
	// The rail offers the menu only on an editable deck (`thumbnail-events`), so it is not gated
	// again here.
	const rail = createSlideRailController({ canEdit: () => true });
	return {
		isSelected: (slideId) => rail.getSnapshot().selectedIds.includes(slideId),
		onClick: (event, slideId, orderedIds) => rail.click(event, slideId, orderedIds),
		openContextMenu(x, y, index, orderedIds) {
			if (!rail.openMenu(x, y, index, orderedIds)) {
				return null;
			}
			const menu = rail.getSnapshot().menu!;
			rail.closeMenu();
			return { ...menu, selectedIndexes: [...menu.selectedIndexes] };
		},
	};
}
