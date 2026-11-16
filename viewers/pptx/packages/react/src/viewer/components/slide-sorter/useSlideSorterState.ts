/**
 * The Slide Sorter's state for React: the shared `createSlideSorterController` read through
 * `useController`, with the overlay's handler shape on top.
 *
 * The keydown listener is registered ONCE, and the controller reads the live props through a
 * ref. Re-subscribing on every render looks harmless and is not: Escape is dispatched to the
 * editor's own window listener first, which sets state; React flushes that render in the
 * browser's microtask checkpoint between listeners, and a re-run effect would remove the
 * sorter's listener before the dispatch reached it (a listener added during a dispatch never
 * receives that event). A stable listener removes the race outright.
 */
import { createSlideSorterController, slideSorterView } from 'ooxml-ui/pptx';
import type { SlideSorterOperation, SlidePasteRequest } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type React from 'react';

import { useController } from '../../hooks/useController';
import type { SlideSectionGroup } from '../../types';
import type { SorterContextMenuState } from './types';

interface UseSlideSorterStateParams {
	slides: PptxSlide[];
	templateElementsBySlideId?: Record<string, PptxElement[]>;
	activeSlideIndex: number;
	canEdit: boolean;
	sectionGroups: SlideSectionGroup[];
	onSelectSlide: (index: number) => void;
	onMoveSlide: (fromIndex: number, toIndex: number) => void;
	onDeleteSlides: (indexes: number[]) => void;
	onDuplicateSlides: (indexes: number[]) => void;
	onPasteSlides: (request: SlidePasteRequest) => void;
	onToggleHideSlides: (indexes: number[]) => void;
	onClose: () => void;
}

// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
export function useSlideSorterState(params: UseSlideSorterStateParams) {
	const { slides, sectionGroups, onSelectSlide, onClose } = params;
	const latest = useRef(params);
	latest.current = params;

	const [controller] = useState(() =>
		createSlideSorterController({
			slides: () => latest.current.slides,
			activeIndex: () => latest.current.activeSlideIndex,
			canEdit: () => latest.current.canEdit,
			clipboardTemplates: () => latest.current.templateElementsBySlideId,
			onOperation: (operation: SlideSorterOperation, indexes) => {
				const list = [...indexes];
				const current = latest.current;
				if (operation === 'delete') {
					current.onDeleteSlides(list);
				} else if (operation === 'duplicate') {
					current.onDuplicateSlides(list);
				} else {
					current.onToggleHideSlides(list);
				}
			},
			onPaste: (request) => latest.current.onPasteSlides(request),
			onMove: (from, to) => latest.current.onMoveSlide(from, to),
			onClose: () => latest.current.onClose(),
		}),
	);
	const state = useController(controller);
	const view = useMemo(() => slideSorterView(state, slides), [state, slides]);
	const backdropRef = useRef<HTMLDivElement>(null);

	useEffect(() => {
		const handleKeyDown = (event: KeyboardEvent): void => {
			controller.handleKeydown(event);
		};
		window.addEventListener('keydown', handleKeyDown);
		return () => window.removeEventListener('keydown', handleKeyDown);
	}, [controller]);

	const setZoom = useCallback(
		(value: number | ((zoom: number) => number)) =>
			controller.setZoom(
				typeof value === 'function' ? value(controller.getSnapshot().sorter.zoom) : value,
			),
		[controller],
	);

	const handlers = useMemo(
		() => ({
			handleSlideClick: (event: React.MouseEvent, index: number) => controller.select(index, event),
			handleContextMenu: (event: React.MouseEvent, index: number) => {
				if (controller.openMenu(index, event.clientX, event.clientY)) {
					event.preventDefault();
					event.stopPropagation();
				}
			},
			closeContextMenu: controller.closeMenu,
			handleDeleteSelected: () => controller.run('delete'),
			handleDuplicateSelected: () => controller.run('duplicate'),
			handleCopySelected: () => controller.run('copy'),
			handlePaste: () => controller.run('paste'),
			handleToggleHideSelected: () => controller.run('toggle-hidden'),
			handleDragStart: (event: React.DragEvent, index: number) => {
				event.dataTransfer.setData('text/plain', String(index));
				event.dataTransfer.effectAllowed = 'move';
				controller.dragStart(index);
			},
			handleDragOver: (event: React.DragEvent, index: number) => {
				event.preventDefault();
				event.dataTransfer.dropEffect = 'move';
				controller.dragOver(index);
			},
			handleDragLeave: controller.dragLeave,
			handleDrop: (event: React.DragEvent, toIndex: number) => {
				event.preventDefault();
				controller.drop(toIndex, parseInt(event.dataTransfer.getData('text/plain'), 10));
			},
		}),
		[controller],
	);

	const handleBackdropClick = useCallback(
		(e: React.MouseEvent) => {
			if (e.target === backdropRef.current) {
				onClose();
			}
		},
		[onClose],
	);
	const handleDoubleClick = useCallback((index: number) => onSelectSlide(index), [onSelectSlide]);
	const isSelected = useCallback((slideId: string) => view.selectedIds.has(slideId), [view]);
	const contextMenu: SorterContextMenuState | null = state.menu
		? { x: state.menu.x, y: state.menu.y, slideIndex: state.menu.index }
		: null;

	return {
		...handlers,
		controller,
		dragOverIndex: state.dragOverIndex,
		selectedSlideIds: state.sorter.selectedIds,
		contextMenu,
		zoom: state.sorter.zoom,
		setZoom,
		clipboardSlideIds: state.sorter.clipboardIds,
		hasClipboard: view.menuContext.hasClipboard,
		backdropRef,
		selectedIndexes: view.selectedIndexes,
		isSelected,
		handleBackdropClick,
		handleDoubleClick,
		zoomScale: view.zoomScale,
		gridCols: view.gridColumns,
		showSectionHeaders: sectionGroups.length > 1,
		hasHiddenInSelection: view.menuContext.hasHiddenInSelection,
		hasVisibleInSelection: view.menuContext.hasVisibleInSelection,
		pasteInsertAt: controller.pasteInsertAt,
	};
}
