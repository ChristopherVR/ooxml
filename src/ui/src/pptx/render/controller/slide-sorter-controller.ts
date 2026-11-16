/**
 * The Slide Sorter view as one state owner for every binding: the tile selection and its
 * anchor, the sorter's own slide clipboard, the zoom, the right-click menu, the tile being
 * dragged and the tile under it, and the keyboard.
 *
 * The decisions were already shared (`slide-sorter-state`, `slide-sorter-keymap`,
 * `slide-sorter-context-menu`); each binding still wrapped them in its own copy of the same
 * state machine (a `runAction` that applied the descriptor and called the deck callbacks, a
 * key handler, a menu opener, drag bookkeeping), and the copies drifted: React opened the menu
 * on a read-only deck, vanilla reordered a read-only deck by dragging, Angular and Svelte had
 * no drag reorder at all, and only vanilla undid inside the sorter. This controller is that
 * wrapper, written once.
 *
 * A binding creates one controller per open sorter, reads `getSnapshot()` through its store
 * adapter, derives the drawn values with {@link slideSorterView} from its own slides, and
 * forwards the gestures. The deck itself stays binding state: the controller reads it through
 * `deps` and reports what to do with it through the callbacks.
 *
 * @module render/controller/slide-sorter-controller
 */
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';

import { isEditorTextInputTarget } from '../editor-key-target';
import type { SlidePasteRequest } from '../slide-operations';
import { buildSlideSorterContextMenuEntries } from '../slide-sorter-context-menu';
import type {
	SlideSorterContextMenuCommandId,
	SlideSorterContextMenuContext,
	SlideSorterContextMenuEntry,
} from '../slide-sorter-context-menu';
import { clampSorterZoom, mapSlideSorterKey, SORTER_ZOOM_STEP } from '../slide-sorter-keymap';
import type { SlideSorterKeyActionName, SlideSorterKeyInput } from '../slide-sorter-keymap';
import {
	applySorterAction,
	createSlideSorterState,
	selectSorterSlide,
	sorterGridColumns,
	sorterMenuContext,
	sorterPasteInsertAt,
	sorterSelectionIndexes,
} from '../slide-sorter-state';
import type { SlideSorterState } from '../slide-sorter-state';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

/** Where the right-click menu is open, and over which tile. */
export interface SlideSorterMenuState {
	readonly x: number;
	readonly y: number;
	readonly index: number;
}

export interface SlideSorterControllerState {
	/** Selection, anchor, sorter clipboard and zoom (the shared `SlideSorterState`). */
	readonly sorter: SlideSorterState;
	/** The right-click menu, or null when it is closed. */
	readonly menu: SlideSorterMenuState | null;
	/** The tile being dragged, or null. */
	readonly dragIndex: number | null;
	/** The tile the dragged one is over (the drop preview), or null. */
	readonly dragOverIndex: number | null;
}

/** A deck operation over the selection; the indexes are in descending order. */
export type SlideSorterOperation = 'delete' | 'duplicate' | 'toggle-hidden';

/** Any command the sorter runs: a key action or a menu entry. */
export type SlideSorterCommand = SlideSorterKeyActionName | SlideSorterContextMenuCommandId;

export interface SlideSorterControllerDeps {
	/** The deck as it is now. */
	slides(): readonly PptxSlide[];
	/** The editor's active slide (the anchor of a fresh sorter, the paste fallback). */
	activeIndex(): number;
	/** Whether the host allows edits: gates the menu, drag reorder and deck-writing keys. */
	canEdit(): boolean;
	/** The slides Copy captures, when they differ from `slides()` (defaults to it). */
	clipboardSourceSlides?(): readonly PptxSlide[] | undefined;
	/** Layout and master elements Copy carries with the slides. */
	clipboardTemplates?(): Record<string, PptxElement[]> | undefined;
	/** Apply `operation` to `indexes` (descending, so one-by-one removal stays correct). */
	onOperation(operation: SlideSorterOperation, indexes: readonly number[]): void;
	/** Insert the copied slides. */
	onPaste(request: SlidePasteRequest): void;
	/** Move one slide (drag and drop, or a move button). */
	onMove(from: number, to: number): void;
	/** Leave the sorter (Escape with nothing to collapse). */
	onClose(): void;
	/** Ctrl+Z inside the sorter, when the host routes history through it. */
	onUndo?(): void;
	/** Ctrl+Y or Ctrl+Shift+Z inside the sorter. */
	onRedo?(): void;
}

export interface SlideSorterControllerOptions {
	/** A sorter state to resume (vanilla keeps the clipboard across a re-open). */
	initial?: SlideSorterState | undefined;
}

/** The part of a keyboard event the controller reads and acts on. */
export interface SlideSorterKeyEvent extends SlideSorterKeyInput {
	readonly target: EventTarget | null;
	preventDefault(): void;
	stopPropagation(): void;
	/** Used for the history keys, so the editor's own listener on the same node cannot undo twice. */
	stopImmediatePropagation?(): void;
}

/** Click modifiers for {@link SlideSorterController.select}. */
export interface SlideSorterClickModifiers {
	readonly ctrlKey?: boolean;
	readonly metaKey?: boolean;
	readonly shiftKey?: boolean;
}

export interface SlideSorterController extends PptxController<SlideSorterControllerState> {
	/** A click on tile `index` (Shift extends, Ctrl/Cmd toggles). */
	select(index: number, modifiers?: SlideSorterClickModifiers): void;
	/**
	 * Start over from `initial`, or from the active slide: closes the menu and ends a drag. A
	 * binding whose deck is not readable at construction (Angular's required inputs) creates the
	 * controller with an empty `initial` and resets it once the inputs are set.
	 */
	reset(initial?: SlideSorterState): void;
	/** Replace the selection with `ids` (the binding's own select-all, for instance). */
	selectIds(ids: readonly string[]): void;
	/**
	 * A right click on tile `index` at (`x`, `y`): selects it (keeping a selection it is part of)
	 * and opens the menu. Returns false on a read-only deck, so the binding leaves the native
	 * menu alone.
	 */
	openMenu(index: number, x: number, y: number): boolean;
	closeMenu(): void;
	/** Run a key or menu command over the selection. */
	run(command: SlideSorterCommand): void;
	/**
	 * A keydown while the sorter is open. Closes the menu, resolves the key against the shared
	 * keymap and runs it; returns whether it was handled (the event is then prevented and
	 * stopped).
	 */
	handleKeydown(event: SlideSorterKeyEvent): boolean;
	/** Set the zoom percentage (clamped to the sorter's range). */
	setZoom(zoom: number): void;
	/** One zoom step in or out. */
	zoomBy(direction: 1 | -1): void;
	/** Move one slide, unless the deck is read-only or it would not move. */
	move(from: number, to: number): void;
	dragStart(index: number): void;
	dragOver(index: number): void;
	dragLeave(): void;
	/**
	 * Drop on tile `index`. The source is the tile `dragStart` recorded, or `fromData` (the
	 * index a binding read back from the drag's data) for a drag that began elsewhere.
	 */
	drop(index: number, fromData?: number): void;
	dragEnd(): void;
	/** The selection for the system clipboard's Copy. */
	selectedIndexes(): number[];
	/** Where the system clipboard's Paste inserts. */
	pasteInsertAt(): number;
}

/** The drawn values of a sorter state over `slides`, for the binding's reactive scope. */
export interface SlideSorterView {
	/** Selected tiles in deck order. */
	readonly selectedIndexes: number[];
	readonly selectedIds: ReadonlySet<string>;
	readonly menuContext: SlideSorterContextMenuContext;
	readonly menuEntries: SlideSorterContextMenuEntry[];
	/** Columns of the tile grid at this zoom. */
	readonly gridColumns: number;
	/** The zoom as a scale factor (1 at 100%). */
	readonly zoomScale: number;
}

export function slideSorterView(
	state: SlideSorterControllerState,
	slides: readonly PptxSlide[],
): SlideSorterView {
	const menuContext = sorterMenuContext(state.sorter, slides);
	return {
		selectedIndexes: sorterSelectionIndexes(state.sorter, slides),
		selectedIds: new Set(state.sorter.selectedIds),
		menuContext,
		menuEntries: buildSlideSorterContextMenuEntries(menuContext),
		gridColumns: sorterGridColumns(state.sorter.zoom),
		zoomScale: state.sorter.zoom / 100,
	};
}

const IDLE_DRAG = { dragIndex: null, dragOverIndex: null } as const;

function historyHandler(
	event: SlideSorterKeyEvent,
	deps: SlideSorterControllerDeps,
): (() => void) | undefined {
	if (isEditorTextInputTarget(event.target) || !(event.ctrlKey || event.metaKey) || event.altKey) {
		return undefined;
	}
	const key = event.key.toLowerCase();
	if (key === 'z') {
		return event.shiftKey ? deps.onRedo : deps.onUndo;
	}
	return key === 'y' ? deps.onRedo : undefined;
}

export function createSlideSorterController(
	deps: SlideSorterControllerDeps,
	options: SlideSorterControllerOptions = {},
): SlideSorterController {
	const store = createControllerStore<SlideSorterControllerState>({
		sorter: options.initial ?? createSlideSorterState(deps.slides(), deps.activeIndex()),
		menu: null,
		...IDLE_DRAG,
	});
	const setSorter = (sorter: SlideSorterState): void =>
		store.update((current) => (current.sorter === sorter ? current : { ...current, sorter }));
	const closeMenu = (): void =>
		store.update((current) => (current.menu ? { ...current, menu: null } : current));
	const selectedIndexes = (): number[] =>
		sorterSelectionIndexes(store.getSnapshot().sorter, deps.slides());

	const run = (command: SlideSorterCommand): void => {
		closeMenu();
		const slides = deps.slides();
		const result = applySorterAction(
			store.getSnapshot().sorter,
			slides,
			command,
			deps.activeIndex(),
			deps.clipboardSourceSlides?.() ?? slides,
			deps.clipboardTemplates?.() ?? {},
		);
		setSorter(result.state);
		if (result.paste) {
			deps.onPaste(result.paste);
		}
		if (result.operation && result.indexes.length > 0) {
			deps.onOperation(result.operation, result.indexes);
		}
		if (result.close) {
			deps.onClose();
		}
	};

	const move = (from: number, to: number): void => {
		const count = deps.slides().length;
		if (!deps.canEdit() || from === to || from < 0 || to < 0 || from >= count || to >= count) {
			return;
		}
		deps.onMove(from, to);
	};

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		select(index, modifiers = {}) {
			setSorter(selectSorterSlide(store.getSnapshot().sorter, deps.slides(), index, modifiers));
		},
		reset(initial) {
			store.set({
				sorter: initial ?? createSlideSorterState(deps.slides(), deps.activeIndex()),
				menu: null,
				...IDLE_DRAG,
			});
		},
		selectIds(ids) {
			setSorter({ ...store.getSnapshot().sorter, selectedIds: [...ids] });
		},
		openMenu(index, x, y) {
			if (!deps.canEdit() || !deps.slides()[index]) {
				return false;
			}
			store.update((current) => ({
				...current,
				sorter: selectSorterSlide(current.sorter, deps.slides(), index, {}, true),
				menu: { x, y, index },
			}));
			return true;
		},
		closeMenu,
		run,
		handleKeydown(event) {
			closeMenu();
			const history = historyHandler(event, deps);
			if (history) {
				event.preventDefault();
				if (event.stopImmediatePropagation) {
					event.stopImmediatePropagation();
				} else {
					event.stopPropagation();
				}
				history();
				return true;
			}
			const { action } = mapSlideSorterKey(event, {
				canEdit: deps.canEdit(),
				hasMultiSelection: selectedIndexes().length > 1,
				isTextInputTarget: isEditorTextInputTarget(event.target),
			});
			if (!action) {
				return false;
			}
			event.preventDefault();
			event.stopPropagation();
			run(action);
			return true;
		},
		setZoom(zoom) {
			const sorter = store.getSnapshot().sorter;
			const next = clampSorterZoom(Number.isFinite(zoom) ? zoom : sorter.zoom);
			if (next !== sorter.zoom) {
				setSorter({ ...sorter, zoom: next });
			}
		},
		zoomBy(direction) {
			const sorter = store.getSnapshot().sorter;
			const next = clampSorterZoom(sorter.zoom + direction * SORTER_ZOOM_STEP);
			if (next !== sorter.zoom) {
				setSorter({ ...sorter, zoom: next });
			}
		},
		move,
		dragStart(index) {
			store.update((current) => ({ ...current, dragIndex: index, dragOverIndex: null }));
		},
		dragOver(index) {
			store.update((current) =>
				current.dragOverIndex === index ? current : { ...current, dragOverIndex: index },
			);
		},
		dragLeave() {
			store.update((current) =>
				current.dragOverIndex === null ? current : { ...current, dragOverIndex: null },
			);
		},
		drop(index, fromData) {
			const from = store.getSnapshot().dragIndex ?? fromData;
			store.update((current) =>
				current.dragIndex === null && current.dragOverIndex === null
					? current
					: { ...current, ...IDLE_DRAG },
			);
			if (from !== undefined && Number.isInteger(from)) {
				move(from, index);
			}
		},
		dragEnd() {
			store.update((current) =>
				current.dragIndex === null && current.dragOverIndex === null
					? current
					: { ...current, ...IDLE_DRAG },
			);
		},
		selectedIndexes,
		pasteInsertAt() {
			return sorterPasteInsertAt(selectedIndexes(), deps.activeIndex(), deps.slides().length);
		},
	};
}
