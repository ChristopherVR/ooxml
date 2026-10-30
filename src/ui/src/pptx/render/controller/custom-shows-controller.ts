/**
 * The Custom Shows dialog (Slide Show > Custom Slide Show > Custom Shows), as a headless
 * controller: the list of shows with New, Edit, Copy and Remove, and for the selected show
 * PowerPoint's two lists (slides in the presentation on the left, slides in the custom show on
 * the right) with Add, Remove and Move Up/Down.
 *
 * It edits a draft of the shows through the core operations in
 * `ooxml-core/pptx/editor/render/custom-show-operations`; nothing reaches the deck until the
 * binding calls {@link CustomShowsController.result} on OK. A binding maps {@link
 * CustomShowsDialogView} onto its template and calls the actions; it keeps no show logic.
 *
 * @module render/controller/custom-shows-controller
 */

import type { PptxCustomShow, PptxSlide } from 'ooxml-core/pptx';
import {
	addSlidesToCustomShow,
	clearCustomShowActions,
	copyCustomShow,
	createCustomShow,
	deleteCustomShow,
	moveCustomShowSlide,
	removeSlidesFromCustomShow,
	renameCustomShow,
} from 'ooxml-core/pptx/editor/render/custom-show-operations';

import { createControllerStore } from './controller';
import type { PptxController } from './controller';

/** The dialog's state; replaced (never mutated) on every action. */
export interface CustomShowsState {
	/** The draft shows. */
	readonly shows: readonly PptxCustomShow[];
	/** Ids of shows removed in this session (their jumps are cleared on OK). */
	readonly removedIds: readonly string[];
	/** The show being edited, or null. */
	readonly selectedId: string | null;
	/** Presentation-list selection, as slide relationship ids. */
	readonly leftSelected: readonly string[];
	/** Show-list selection, as positions in the selected show. */
	readonly rightSelected: readonly number[];
}

/** One slide row of either list. */
export interface CustomShowsSlideRow {
	readonly rId: string;
	/** 1-based number in the presentation. */
	readonly number: number;
	/** The slide's title text, when it has one. */
	readonly title: string;
	readonly selected: boolean;
	/** Position in the show list (right list only). */
	readonly position?: number;
}

/** Everything a template needs to draw the dialog. */
export interface CustomShowsDialogView {
	readonly shows: readonly { id: string; name: string; slideCount: number; selected: boolean }[];
	readonly selected: { id: string; name: string } | null;
	readonly presentationSlides: readonly CustomShowsSlideRow[];
	readonly showSlides: readonly CustomShowsSlideRow[];
	readonly canEdit: boolean;
	readonly canAdd: boolean;
	readonly canRemoveSlides: boolean;
	readonly canMoveUp: boolean;
	readonly canMoveDown: boolean;
}

/** What OK commits. */
export interface CustomShowsResult {
	readonly shows: PptxCustomShow[];
	/** The slides with jumps into removed shows cleared, or undefined when none changed. */
	readonly slides: PptxSlide[] | undefined;
	readonly removedIds: readonly string[];
}

export interface CustomShowsController extends PptxController<CustomShowsState> {
	/** Draft `New`: a show with the given name and no slides, selected for editing. */
	create(name: string): void;
	/** Draft `Copy` of the selected show. */
	copySelected(name?: string): void;
	/** Draft `Remove` of the selected show. */
	removeSelected(): void;
	rename(name: string): void;
	select(id: string | null): void;
	toggleLeft(rId: string): void;
	toggleRight(position: number): void;
	/** `Add >>`: append the selected presentation slides (in deck order, from `slides`) to the show. */
	addSelected(slides: readonly PptxSlide[]): void;
	/** Add one slide regardless of the selection (double-click). */
	addSlide(rId: string): void;
	/** `Remove`: drop the selected entries from the show. */
	removeSelectedSlides(): void;
	moveUp(): void;
	moveDown(): void;
	/** The view for `slides` (read live, so edits to the deck show). */
	view(slides: readonly PptxSlide[]): CustomShowsDialogView;
	/** The draft to commit, with jumps into removed shows cleared from `slides`. */
	result(slides: readonly PptxSlide[]): CustomShowsResult;
}

function slideTitle(slide: PptxSlide): string {
	for (const element of slide.elements) {
		const text = 'text' in element && typeof element.text === 'string' ? element.text.trim() : '';
		if (text) {
			return text.split(/\r?\n/)[0]!.slice(0, 60);
		}
	}
	return '';
}

/** Create the dialog controller over the deck's current shows. */
export function createCustomShowsController(
	initial: readonly PptxCustomShow[],
	selectedId: string | null = null,
): CustomShowsController {
	const first = initial.find((s) => s.id === selectedId) ?? initial[0];
	const store = createControllerStore<CustomShowsState>({
		shows: initial.map((s) => ({ ...s, slideRIds: [...s.slideRIds] })),
		removedIds: [],
		selectedId: first?.id ?? null,
		leftSelected: [],
		rightSelected: [],
	});
	const clearSelection = { leftSelected: [], rightSelected: [] } as const;

	const edit = (
		recipe: (
			state: CustomShowsState,
			selected: PptxCustomShow,
		) => Partial<CustomShowsState> | undefined,
	): void =>
		store.update((state) => {
			const selected = state.shows.find((s) => s.id === state.selectedId);
			const patch = selected ? recipe(state, selected) : undefined;
			return patch ? { ...state, ...patch } : state;
		});

	const move = (direction: -1 | 1): void =>
		edit((state, show) => {
			const occupied = new Set(state.rightSelected);
			let shows = state.shows;
			// Walk from the leading edge so a selected block moves as one; an entry blocked by
			// the list end or another selected entry stays put.
			for (const position of [...occupied].sort((x, y) => direction * (y - x))) {
				const target = position + direction;
				if (target < 0 || target >= show.slideRIds.length || occupied.has(target)) {
					continue;
				}
				shows = moveCustomShowSlide(shows, show.id, position, target).shows;
				occupied.delete(position);
				occupied.add(target);
			}
			return { shows, rightSelected: [...occupied].sort((x, y) => x - y) };
		});

	const addSlides = (rIds: readonly string[]): void =>
		edit((state, show) =>
			rIds.length === 0
				? undefined
				: {
						shows: addSlidesToCustomShow(state.shows, show.id, rIds).shows,
						...clearSelection,
					},
		);

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		create(name) {
			store.update((state) => {
				const r = createCustomShow(state.shows, name);
				return { ...state, shows: r.shows, selectedId: r.show?.id ?? null, ...clearSelection };
			});
		},
		copySelected(name) {
			store.update((state) => {
				const r = state.selectedId
					? copyCustomShow(state.shows, state.selectedId, name)
					: undefined;
				return r?.show
					? { ...state, shows: r.shows, selectedId: r.show.id, ...clearSelection }
					: state;
			});
		},
		removeSelected() {
			store.update((state) => {
				if (!state.selectedId) {
					return state;
				}
				const index = state.shows.findIndex((s) => s.id === state.selectedId);
				const r = deleteCustomShow(state.shows, state.selectedId);
				const neighbour = r.shows[Math.min(index, r.shows.length - 1)];
				return {
					...state,
					shows: r.shows,
					removedIds: [...state.removedIds, state.selectedId],
					selectedId: neighbour?.id ?? null,
					...clearSelection,
				};
			});
		},
		rename(name) {
			edit((state, show) => ({ shows: renameCustomShow(state.shows, show.id, name).shows }));
		},
		select(id) {
			store.update((state) =>
				state.selectedId === id ? state : { ...state, selectedId: id, ...clearSelection },
			);
		},
		toggleLeft(rId) {
			store.update((state) => ({
				...state,
				leftSelected: state.leftSelected.includes(rId)
					? state.leftSelected.filter((r) => r !== rId)
					: [...state.leftSelected, rId],
			}));
		},
		toggleRight(position) {
			store.update((state) => ({
				...state,
				rightSelected: state.rightSelected.includes(position)
					? state.rightSelected.filter((p) => p !== position)
					: [...state.rightSelected, position],
			}));
		},
		addSelected(slides) {
			const chosen = new Set(store.getSnapshot().leftSelected);
			addSlides(slides.filter((slide) => chosen.has(slide.rId)).map((slide) => slide.rId));
		},
		addSlide: (rId) => addSlides([rId]),
		removeSelectedSlides() {
			edit((state, show) =>
				state.rightSelected.length === 0
					? undefined
					: {
							shows: removeSlidesFromCustomShow(state.shows, show.id, state.rightSelected).shows,
							rightSelected: [],
						},
			);
		},
		moveUp: () => move(-1),
		moveDown: () => move(1),
		view(slides) {
			const state = store.getSnapshot();
			const selected = state.shows.find((s) => s.id === state.selectedId);
			const byRId = new Map(slides.map((slide, i) => [slide.rId, { slide, number: i + 1 }]));
			const row = (rId: string, selectedRow: boolean, position?: number) => {
				const entry = byRId.get(rId);
				return {
					rId,
					number: entry?.number ?? 0,
					title: entry ? slideTitle(entry.slide) : '',
					selected: selectedRow,
					...(position === undefined ? {} : { position }),
				};
			};
			const showRows = (selected?.slideRIds ?? []).flatMap((rId, position) =>
				byRId.has(rId) ? [row(rId, state.rightSelected.includes(position), position)] : [],
			);
			const positions = state.rightSelected;
			const last = (selected?.slideRIds.length ?? 0) - 1;
			return {
				shows: state.shows.map((s) => ({
					id: s.id,
					name: s.name,
					slideCount: s.slideRIds.length,
					selected: s.id === state.selectedId,
				})),
				selected: selected ? { id: selected.id, name: selected.name } : null,
				presentationSlides: slides.map((slide) =>
					row(slide.rId, state.leftSelected.includes(slide.rId)),
				),
				showSlides: showRows,
				canEdit: selected !== undefined,
				canAdd: selected !== undefined && state.leftSelected.length > 0,
				canRemoveSlides: positions.length > 0,
				canMoveUp: positions.length > 0 && Math.min(...positions) > 0,
				canMoveDown: positions.length > 0 && Math.max(...positions) < last,
			};
		},
		result(slides) {
			const state = store.getSnapshot();
			let current: PptxSlide[] = [...slides];
			let changed = false;
			const kept = new Set(state.shows.map((show) => show.id));
			for (const id of state.removedIds.filter((removed) => !kept.has(removed))) {
				const cleared = clearCustomShowActions(current, id);
				if (cleared.cleared > 0) {
					current = cleared.slides;
					changed = true;
				}
			}
			return {
				shows: state.shows.map((s) => ({ ...s, slideRIds: [...s.slideRIds] })),
				slides: changed ? current : undefined,
				removedIds: state.removedIds,
			};
		},
	};
}
