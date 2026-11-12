/**
 * PowerPoint's Ribbon Display Options as one state owner for every binding: the mode (Always
 * show Ribbon, Show tabs only, Full-screen mode), the temporary peek that lays the commands
 * over the slide, the menu, and Options > General > "Collapse the ribbon automatically".
 *
 * A binding keeps one controller per viewer (it must outlive a ribbon rebuild), reads
 * `getSnapshot()` through its store adapter and maps `layout` onto its template:
 *
 *  - `layout.titleBar` false: hide the title bar;
 *  - `layout.tabs` false: hide the tab row and the command row above it;
 *  - `layout.commands`: `docked` (in the layout), `overlay` (absolutely positioned over the
 *    slide, below the tab row) or `hidden`;
 *  - `layout.reveal`: draw the slim strip that calls `reveal()`.
 *
 * It forwards four user gestures: a tab click (`tabChosen`), a double click on a tab
 * (`toggleCollapsed`), a press outside the ribbon (`dismiss`) and key presses
 * (`handleKeydown`: Ctrl+F1 and Escape).
 *
 * @module render/controller/ribbon-display-controller
 */

import {
	normalizeRibbonDisplayMode,
	ribbonDisplayLayout,
	ribbonDisplayMenu,
	toggleRibbonCollapsed,
} from '../ribbon-display';
import type {
	RibbonDisplayLayout,
	RibbonDisplayMenuItem,
	RibbonDisplayMode,
} from '../ribbon-display';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

/**
 * Viewports this short collapse the ribbon to its tab row while "Collapse the ribbon
 * automatically" is on, the same height the Word and Excel ribbons use.
 */
export const RIBBON_AUTO_COLLAPSE_MEDIA = '(max-height: 500px)';

export interface RibbonDisplayState {
	/** The mode the user or the host chose. */
	readonly mode: RibbonDisplayMode;
	/**
	 * The mode drawn. It differs from `mode` only while the ribbon is collapsed automatically:
	 * then `mode` is `always` and this is `tabs`.
	 */
	readonly effectiveMode: RibbonDisplayMode;
	/** The commands are called up temporarily (a tab in `tabs`, the strip in `fullScreen`). */
	readonly peek: boolean;
	/** The Ribbon Display Options menu is open. */
	readonly menuOpen: boolean;
	/** What to draw for `effectiveMode` and `peek`. */
	readonly layout: RibbonDisplayLayout;
	/** The menu's radio items, the drawn mode checked. */
	readonly menu: readonly RibbonDisplayMenuItem[];
}

export interface RibbonDisplayControllerOptions {
	/** The host's `ribbonDisplay` value; anything unknown is `always`. */
	initial?: unknown;
	/** The user picked a mode (menu, Ctrl+F1, double click on a tab). Not called for `setMode`. */
	onChange?: ((mode: RibbonDisplayMode) => void) | undefined;
}

/** The part of a keyboard event the controller reads. */
export interface RibbonDisplayKey {
	readonly key: string;
	readonly ctrlKey: boolean;
	preventDefault(): void;
}

export interface RibbonDisplayController extends PptxController<RibbonDisplayState> {
	/** Set the mode from the host (an option, a prop, a stored value); `onChange` is not called. */
	setMode(mode: unknown): void;
	/** The user picked `mode` in the menu: applies it, closes the menu and calls `onChange`. */
	choose(mode: RibbonDisplayMode): void;
	/** Ctrl+F1 and a double click on a tab: switch between the docked ribbon and the tab row. */
	toggleCollapsed(): void;
	/** A tab was chosen; pass its id. `file` opens the backstage and never peeks. */
	tabChosen(tab: string): void;
	/** The strip of full-screen mode was pressed: bring the ribbon back over the slide. */
	reveal(): void;
	/** A press outside the ribbon, or Escape: put the peeked commands and the menu away. */
	dismiss(): void;
	setMenuOpen(open: boolean): void;
	toggleMenu(): void;
	/** Options > General > "Collapse the ribbon automatically". */
	setAutoCollapse(enabled: boolean): void;
	/** Whether the viewport matches {@link RIBBON_AUTO_COLLAPSE_MEDIA}. */
	setShortViewport(short: boolean): void;
	/** Ctrl+F1 toggles, Escape dismisses. Returns true when the key was used. */
	handleKeydown(event: RibbonDisplayKey): boolean;
}

interface Inputs {
	mode: RibbonDisplayMode;
	peek: boolean;
	menuOpen: boolean;
	autoCollapse: boolean;
	shortViewport: boolean;
	/** The user picked a mode themselves, which wins over the automatic collapse. */
	chosen: boolean;
}

function derive(inputs: Inputs): RibbonDisplayState {
	const automatic =
		inputs.mode === 'always' && inputs.autoCollapse && inputs.shortViewport && !inputs.chosen;
	const effectiveMode: RibbonDisplayMode = automatic ? 'tabs' : inputs.mode;
	// Nothing can be peeked while the commands are docked.
	const peek = effectiveMode === 'always' ? false : inputs.peek;
	return {
		mode: inputs.mode,
		effectiveMode,
		peek,
		menuOpen: inputs.menuOpen,
		layout: ribbonDisplayLayout(effectiveMode, peek),
		menu: ribbonDisplayMenu(effectiveMode),
	};
}

export function createRibbonDisplayController(
	options: RibbonDisplayControllerOptions = {},
): RibbonDisplayController {
	let inputs: Inputs = {
		mode: normalizeRibbonDisplayMode(options.initial),
		peek: false,
		menuOpen: false,
		autoCollapse: false,
		shortViewport: false,
		chosen: false,
	};
	const store = createControllerStore(derive(inputs));
	const commit = (patch: Partial<Inputs>): void => {
		const next = { ...inputs, ...patch };
		if ((Object.keys(next) as (keyof Inputs)[]).every((key) => next[key] === inputs[key])) {
			return;
		}
		inputs = next;
		store.set(derive(inputs));
	};
	const choose = (mode: RibbonDisplayMode): void => {
		const changed = mode !== store.getSnapshot().effectiveMode;
		commit({ mode, chosen: true, peek: false, menuOpen: false });
		if (changed) {
			options.onChange?.(mode);
		}
	};
	const dismiss = (): void => commit({ peek: false, menuOpen: false });

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		setMode: (mode) =>
			commit({ mode: normalizeRibbonDisplayMode(mode), peek: false, menuOpen: false }),
		choose,
		toggleCollapsed: () => choose(toggleRibbonCollapsed(store.getSnapshot().effectiveMode)),
		tabChosen(tab) {
			if (tab !== 'file' && store.getSnapshot().effectiveMode === 'tabs') {
				commit({ peek: true });
			}
		},
		reveal() {
			if (store.getSnapshot().effectiveMode === 'fullScreen') {
				commit({ peek: true });
			}
		},
		dismiss,
		setMenuOpen: (open) => commit({ menuOpen: open }),
		toggleMenu: () => commit({ menuOpen: !inputs.menuOpen }),
		// Changing the setting is a fresh instruction: it applies again even after the user
		// picked a mode by hand.
		setAutoCollapse: (enabled) =>
			commit({
				autoCollapse: enabled,
				chosen: enabled === inputs.autoCollapse ? inputs.chosen : false,
			}),
		setShortViewport: (short) => commit({ shortViewport: short }),
		handleKeydown(event) {
			if (event.key === 'F1' && event.ctrlKey) {
				event.preventDefault();
				choose(toggleRibbonCollapsed(store.getSnapshot().effectiveMode));
				return true;
			}
			const { peek, menuOpen } = store.getSnapshot();
			if (event.key === 'Escape' && (peek || menuOpen)) {
				dismiss();
				return true;
			}
			return false;
		},
	};
}
