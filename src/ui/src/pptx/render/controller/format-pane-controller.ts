/**
 * The Format pane's open sections and its launch requests, one controller per
 * editor instance (never module state: test workers share modules).
 *
 * Which sections are open is kept per selection category and section
 * (`text:font`), in memory only: closing Fill on a shape keeps it closed for
 * the next shape, and a reload starts from the defaults again. Nothing is
 * persisted.
 *
 * Launchers (the context menu's Format Shape, Size and Position, Edit Alt Text,
 * a ribbon group's dialog launcher) call {@link FormatPaneController.requestFormatPane}.
 * The controller asks the binding to show the Properties tab (its `showPane`
 * dependency) and keeps the request until the pane takes it with
 * {@link FormatPaneController.takeRequest} after rendering, which opens the
 * section and tells the pane what to reveal and focus.
 *
 * @module render/controller/format-pane-controller
 */
import type {
	FormatSectionId,
	FormatSectionView,
	FormatSelectionCategory,
} from '../inspector/format-sections';
import { createControllerStore } from './controller';
import type { PptxController } from './controller';

/** A pending "show this section" request. */
export interface FormatPaneRequest {
	/** The section to reveal (`null`: the first section of the selection). */
	readonly section: FormatSectionId | null;
	/** Move keyboard focus to the section heading. */
	readonly focus: boolean;
	/** Increases with every request, so a repeat of the same request is seen. */
	readonly seq: number;
}

/** What the pane has to do once it has rendered a request. */
export interface FormatPaneReveal {
	readonly section: FormatSectionId;
	readonly focus: boolean;
}

export interface FormatPaneState {
	/** User choices, keyed `category:section`. Absent keys use the section default. */
	readonly open: Readonly<Record<string, boolean>>;
	readonly request: FormatPaneRequest | null;
}

export interface FormatPaneControllerDeps {
	/** Open the inspector and switch it to the Properties tab. */
	showPane?: () => void;
}

export interface FormatPaneController extends PptxController<FormatPaneState> {
	/** Whether `view` is open for `category` (the user's choice, else its default). */
	isOpen(category: FormatSelectionCategory, view: FormatSectionView): boolean;
	setOpen(category: FormatSelectionCategory, section: FormatSectionId, open: boolean): void;
	/** Show the Properties tab and, once rendered, expand and reveal `section`. */
	requestFormatPane(section?: FormatSectionId | null, options?: { focus?: boolean }): void;
	/**
	 * Take the pending request for what the pane now shows: opens the section
	 * (the first one when the requested section is not shown) and returns what to
	 * reveal, or `null` when there is no request or nothing to show.
	 */
	takeRequest(
		category: FormatSelectionCategory,
		sections: readonly FormatSectionView[],
	): FormatPaneReveal | null;
	/**
	 * Record the category the pane shows; `true` when it changed since the last
	 * call (the pane then scrolls to the top, unless a request is pending).
	 */
	noteCategory(category: FormatSelectionCategory): boolean;
	/** Replace the `showPane` dependency (bindings that create the controller early). */
	setShowPane(showPane: (() => void) | undefined): void;
}

const key = (category: FormatSelectionCategory, section: FormatSectionId): string =>
	`${category}:${section}`;

/** Pure read of the open state, for bindings that derive it in their own reactive scope. */
export function formatSectionOpen(
	state: FormatPaneState,
	category: FormatSelectionCategory,
	view: FormatSectionView,
): boolean {
	return state.open[key(category, view.id)] ?? view.defaultOpen;
}

export function createFormatPaneController(
	deps: FormatPaneControllerDeps = {},
): FormatPaneController {
	const store = createControllerStore<FormatPaneState>({ open: {}, request: null });
	let showPane = deps.showPane;
	let seq = 0;
	let lastCategory: FormatSelectionCategory | null = null;

	const setOpen = (
		category: FormatSelectionCategory,
		section: FormatSectionId,
		open: boolean,
	): void => {
		store.update((state) =>
			state.open[key(category, section)] === open
				? state
				: { ...state, open: { ...state.open, [key(category, section)]: open } },
		);
	};

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		isOpen: (category, view) => formatSectionOpen(store.getSnapshot(), category, view),
		setOpen,
		requestFormatPane(section = null, options = {}) {
			seq += 1;
			store.update((state) => ({
				...state,
				request: { section, focus: options.focus ?? false, seq },
			}));
			showPane?.();
		},
		takeRequest(category, sections) {
			const request = store.getSnapshot().request;
			if (!request) {
				return null;
			}
			const target = sections.find((view) => view.id === request.section) ?? sections[0] ?? null;
			store.update((state) => ({ ...state, request: null }));
			if (!target) {
				return null;
			}
			setOpen(category, target.id, true);
			return { section: target.id, focus: request.focus };
		},
		noteCategory(category) {
			const changed = lastCategory !== null && lastCategory !== category;
			lastCategory = category;
			return changed;
		},
		setShowPane(next) {
			showPane = next;
		},
	};
}
