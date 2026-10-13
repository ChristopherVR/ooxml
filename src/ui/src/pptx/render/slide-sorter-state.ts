import { slideSorterPasteIndexes } from './slide-sorter-context-menu';
import type { SlideSorterContextMenuContext } from './slide-sorter-context-menu';
import { clampSorterZoom, SORTER_ZOOM_STEP } from './slide-sorter-keymap';
import type { SlideSorterKeyActionName } from './slide-sorter-keymap';
import {
	captureSlideClipboard,
	createSlidePaste,
	captureSlideClipboardTemplates,
} from './slide-operations';
import type { SlidePasteRequest } from './slide-operations';
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';

export interface SorterSlide {
	id: string;
	hidden?: boolean;
}
export interface SlideSorterState {
	selectedIds: string[];
	anchor: number;
	clipboardIds: string[];
	clipboardSlides?: PptxSlide[];
	clipboardTemplates?: Record<string, PptxElement[]>;
	zoom: number;
}
export function createSlideSorterState(
	slides: readonly SorterSlide[],
	active: number,
): SlideSorterState {
	return {
		selectedIds: slides[active] ? [slides[active].id] : [],
		anchor: active,
		clipboardIds: [],
		zoom: 100,
	};
}
export function selectSorterSlide(
	state: SlideSorterState,
	slides: readonly SorterSlide[],
	index: number,
	modifiers: { ctrlKey?: boolean; metaKey?: boolean; shiftKey?: boolean } = {},
	context = false,
): SlideSorterState {
	const slide = slides[index];
	if (!slide || (context && state.selectedIds.includes(slide.id))) {
		return state;
	}
	if (!context && modifiers.shiftKey) {
		const range = slides
			.slice(Math.min(state.anchor, index), Math.max(state.anchor, index) + 1)
			.map((s) => s.id);
		return {
			...state,
			selectedIds:
				modifiers.ctrlKey || modifiers.metaKey
					? [...new Set([...state.selectedIds, ...range])]
					: range,
		};
	}
	if (!context && (modifiers.ctrlKey || modifiers.metaKey)) {
		return {
			...state,
			anchor: index,
			selectedIds: state.selectedIds.includes(slide.id)
				? state.selectedIds.filter((id) => id !== slide.id)
				: [...state.selectedIds, slide.id],
		};
	}
	return { ...state, anchor: index, selectedIds: [slide.id] };
}
export function sorterSelectionIndexes(
	state: SlideSorterState,
	slides: readonly SorterSlide[],
): number[] {
	return slideSorterPasteIndexes(state.selectedIds, slides);
}
export function sorterMenuContext(
	state: SlideSorterState,
	slides: readonly SorterSlide[],
): SlideSorterContextMenuContext {
	const indexes = sorterSelectionIndexes(state, slides);
	return {
		selectedCount: indexes.length,
		hasClipboard: Boolean(
			state.clipboardSlides?.length || slideSorterPasteIndexes(state.clipboardIds, slides).length,
		),
		hasHiddenInSelection: indexes.some((i) => slides[i].hidden),
		hasVisibleInSelection: indexes.some((i) => !slides[i].hidden),
		wouldDeleteAllSlides: indexes.length >= slides.length,
	};
}
export function sorterGridColumns(zoom: number): number {
	return zoom >= 180 ? 2 : zoom >= 140 ? 3 : zoom >= 100 ? 4 : zoom >= 70 ? 5 : 6;
}
/** A decision descriptor; adapters apply deck operations in descending index order. */
export function applySorterAction(
	state: SlideSorterState,
	slides: readonly SorterSlide[],
	action: SlideSorterKeyActionName | 'toggle-hidden',
	active: number,
	clipboardSourceSlides?: readonly PptxSlide[],
	clipboardTemplates: Record<string, PptxElement[]> = {},
): {
	state: SlideSorterState;
	indexes: number[];
	operation?: 'duplicate' | 'delete' | 'toggle-hidden';
	close?: boolean;
	paste?: SlidePasteRequest;
} {
	const indexes = sorterSelectionIndexes(state, slides);
	switch (action) {
		case 'close':
			return { state, indexes: [], close: true };
		case 'collapseSelection':
			return {
				state: { ...state, selectedIds: slides[active] ? [slides[active].id] : [], anchor: active },
				indexes: [],
			};
		case 'selectAll':
			return { state: { ...state, selectedIds: slides.map((s) => s.id) }, indexes: [] };
		case 'copy':
			return {
				state: {
					...state,
					clipboardIds: indexes.map((i) => slides[i].id),
					clipboardSlides: captureSlideClipboard(
						clipboardSourceSlides ?? (slides as readonly PptxSlide[]),
						indexes.map((i) => slides[i].id),
					),
					clipboardTemplates: captureSlideClipboardTemplates(
						clipboardTemplates,
						indexes.map((i) => slides[i].id),
					),
				},
				indexes: [],
			};
		case 'zoomIn':
		case 'zoomOut':
			return {
				state: {
					...state,
					zoom: clampSorterZoom(
						state.zoom + (action === 'zoomIn' ? SORTER_ZOOM_STEP : -SORTER_ZOOM_STEP),
					),
				},
				indexes: [],
			};
		case 'paste':
			if (!state.clipboardSlides?.length) return { state, indexes: [] };
			{
				const insertAt = sorterPasteInsertAt(indexes, active, slides.length);
				const paste = createSlidePaste(state.clipboardSlides, insertAt, state.clipboardTemplates);
				return {
					state: { ...state, selectedIds: paste.copies.map((slide) => slide.id), anchor: insertAt },
					indexes: [],
					paste,
				};
			}
		case 'delete':
			return {
				state: indexes.length < slides.length ? { ...state, selectedIds: [] } : state,
				indexes: indexes.length < slides.length ? indexes.reverse() : [],
				operation: 'delete',
			};
		case 'duplicate':
		case 'toggle-hidden':
			return { state, indexes: indexes.reverse(), operation: action };
	}
}

/** The destination follows the last selected tile, or the active slide. */
export function sorterPasteInsertAt(
	indexes: readonly number[],
	active: number,
	slideCount: number,
): number {
	return Math.min(slideCount, Math.max(0, (indexes.at(-1) ?? active) + 1));
}
