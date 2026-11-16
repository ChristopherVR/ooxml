import { createSlideRailController } from 'ooxml-ui/pptx';
import { useCallback, useEffect, useMemo, useState } from 'react';

import { useController } from '../../hooks/useController';
import type { SectionContextMenuState, SlideContextMenuState } from './types';

// ---------------------------------------------------------------------------
// Hook return type
// ---------------------------------------------------------------------------

export interface SlidePaneCallbacks {
	collapsedSections: Record<string, boolean>;
	renamingSectionId: string | null;
	renameValue: string;
	sectionContextMenu: SectionContextMenuState | null;
	slideCtxMenu: SlideContextMenuState | null;
	setRenameValue: (value: string) => void;
	handleDragStart: (e: React.DragEvent, slideIndex: number) => void;
	handleDragOver: (e: React.DragEvent) => void;
	handleDrop: (e: React.DragEvent, toIndex: number) => void;
	/**
	 * Collapse/expand a section. `isCollapsed` is the CURRENT effective state
	 * (which may come from the model's `collapsed` flag rather than this hook's
	 * override map), so the flip cannot disagree with what is on screen.
	 */
	toggleSection: (sectionId: string, isCollapsed: boolean) => void;
	startRename: (sectionId: string, currentLabel: string) => void;
	commitRename: () => void;
	cancelRename: () => void;
	handleSectionContextMenu: (
		e: React.MouseEvent,
		sectionId: string,
		sectionIndex: number,
		totalSections: number,
	) => void;
	/**
	 * A right-click on thumbnail `slideIndex`: the menu acts on the whole selection when the slide
	 * is part of it, else on that slide (the shared slide rail controller).
	 */
	handleOpenSlideCtxMenu: (
		x: number,
		y: number,
		slideIndex: number,
		orderedIds: readonly string[],
	) => void;
	closeSectionContextMenu: () => void;
	closeSlideCtxMenu: () => void;
	/** Ctrl/Cmd/Shift-aware multi-select; ids because indexes shift under insert/delete/move. */
	selectedSlideIds: string[];
	/**
	 * Resolve one click on a slide thumbnail: updates the multi-selection and
	 * always reports the plain "make this slide active" index too, since a
	 * click (even a Ctrl/Shift one) also moves the canvas to the slide clicked,
	 * matching the sorter overlay and PowerPoint itself.
	 */
	handleSlideClick: (
		event: { ctrlKey: boolean; metaKey: boolean; shiftKey: boolean },
		slideId: string,
		orderedIds: readonly string[],
	) => void;
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useSlidePaneCallbacks(
	onMoveSlide: (fromIndex: number, toIndex: number) => void,
	onRenameSection?: (sectionId: string, newName: string) => void,
	onToggleSectionCollapse?: (sectionId: string) => void,
): SlidePaneCallbacks {
	const [collapsedSections, setCollapsedSections] = useState<Record<string, boolean>>({});
	const [renamingSectionId, setRenamingSectionId] = useState<string | null>(null);
	const [renameValue, setRenameValue] = useState('');
	const [sectionContextMenu, setSectionContextMenu] = useState<SectionContextMenuState | null>(
		null,
	);
	// The thumbnail multi-selection and its menu: the shared slide rail controller. The pane only
	// opens the menu on an editable deck (`SlideItem`), so it is not gated again here.
	const [rail] = useState(() => createSlideRailController({ canEdit: () => true }));
	const railState = useController(rail);
	const selectedSlideIds = railState.selectedIds as string[];
	const slideCtxMenu = useMemo<SlideContextMenuState | null>(
		() =>
			railState.menu
				? {
						x: railState.menu.x,
						y: railState.menu.y,
						slideIndex: railState.menu.index,
						selectedIndexes: [...railState.menu.selectedIndexes],
					}
				: null,
		[railState.menu],
	);

	const handleSlideClick = useCallback(
		(
			event: { ctrlKey: boolean; metaKey: boolean; shiftKey: boolean },
			slideId: string,
			orderedIds: readonly string[],
		) => rail.click(event, slideId, orderedIds),
		[rail],
	);

	// Close context menus on outside click
	useEffect(() => {
		if (!sectionContextMenu && !slideCtxMenu) {
			return;
		}
		const handler = () => {
			setSectionContextMenu(null);
			rail.closeMenu();
		};
		document.addEventListener('click', handler);
		return () => document.removeEventListener('click', handler);
	}, [sectionContextMenu, slideCtxMenu, rail]);

	// ── Drag handlers ──
	const handleDragStart = useCallback((e: React.DragEvent, slideIndex: number) => {
		e.dataTransfer.setData('text/plain', String(slideIndex));
		e.dataTransfer.effectAllowed = 'move';
	}, []);

	const handleDragOver = useCallback((e: React.DragEvent) => {
		e.preventDefault();
		e.dataTransfer.dropEffect = 'move';
	}, []);

	const handleDrop = useCallback(
		(e: React.DragEvent, toIndex: number) => {
			e.preventDefault();
			const fromStr = e.dataTransfer.getData('text/plain');
			const fromIndex = parseInt(fromStr, 10);
			if (!isNaN(fromIndex) && fromIndex !== toIndex) {
				onMoveSlide(fromIndex, toIndex);
			}
		},
		[onMoveSlide],
	);

	const toggleSection = useCallback(
		(sectionId: string, isCollapsed: boolean) => {
			setCollapsedSections((prev) => ({
				...prev,
				[sectionId]: !isCollapsed,
			}));
			// Write the flag back to the model too. Vue, Angular, Svelte and
			// Vanilla all do, so React was the only binding whose collapsed
			// sections were forgotten on save.
			onToggleSectionCollapse?.(sectionId);
		},
		[onToggleSectionCollapse],
	);

	// ── Rename handlers ──
	const startRename = useCallback((sectionId: string, currentLabel: string) => {
		setRenamingSectionId(sectionId);
		setRenameValue(currentLabel);
		setSectionContextMenu(null);
	}, []);

	const commitRename = useCallback(() => {
		if (renamingSectionId && renameValue.trim().length > 0) {
			onRenameSection?.(renamingSectionId, renameValue.trim());
		}
		setRenamingSectionId(null);
		setRenameValue('');
	}, [renamingSectionId, renameValue, onRenameSection]);

	const cancelRename = useCallback(() => {
		setRenamingSectionId(null);
		setRenameValue('');
	}, []);

	// ── Section context menu handler ──
	const handleSectionContextMenu = useCallback(
		(e: React.MouseEvent, sectionId: string, sectionIndex: number, totalSections: number) => {
			e.preventDefault();
			e.stopPropagation();
			setSectionContextMenu({
				x: e.clientX,
				y: e.clientY,
				sectionId,
				sectionIndex,
				totalSections,
			});
		},
		[],
	);

	const handleOpenSlideCtxMenu = useCallback(
		(x: number, y: number, slideIndex: number, orderedIds: readonly string[]) => {
			rail.openMenu(x, y, slideIndex, orderedIds);
		},
		[rail],
	);

	const closeSectionContextMenu = useCallback(() => {
		setSectionContextMenu(null);
	}, []);

	const closeSlideCtxMenu = rail.closeMenu;

	return {
		collapsedSections,
		renamingSectionId,
		renameValue,
		sectionContextMenu,
		slideCtxMenu,
		setRenameValue,
		handleDragStart,
		handleDragOver,
		handleDrop,
		toggleSection,
		startRename,
		commitRename,
		cancelRename,
		handleSectionContextMenu,
		handleOpenSlideCtxMenu,
		closeSectionContextMenu,
		closeSlideCtxMenu,
		selectedSlideIds,
		handleSlideClick,
	};
}
