/**
 * The Find and Replace bar as one state owner for every binding: whether the bar is open and in
 * which mode (Find, Ctrl+F, or Find and Replace, Ctrl+H), and the shared search session behind it
 * (`find-replace-session`).
 *
 * The session was already shared; the bar's open state was not, and drifted: React toggled its bar
 * from a window listener of its own and kept the old matches' query without searching again on
 * open, Vue and Svelte searched again on open, Angular kept two separate bars with two flags, and
 * vanilla's open state lived in its ribbon. Now: Ctrl+F and Ctrl+H toggle the bar (the other
 * chord switches the mode of an open bar instead of closing it), opening searches the kept query
 * again from the first match, and closing drops the matches (and their canvas highlights) while
 * keeping the query, replacement and scopes for the next open.
 *
 * A binding keeps one controller per viewer, reads `getSnapshot()` through its store adapter,
 * routes the editor keymap's `find` / `findReplace` actions to {@link FindReplaceController.toggle}
 * and calls {@link FindReplaceController.slidesChanged} when its deck changes (a replace searches
 * again once the replaced slides land).
 *
 * @module render/controller/find-replace-controller
 */
import { findReplaceSessionActions } from '../find-replace-session-actions';
import type {
	FindReplaceSessionActions,
	FindReplaceSessionHost,
} from '../find-replace-session-actions';
import { INITIAL_FIND_REPLACE_SESSION } from '../find-replace-session';
import type { FindReplaceSession } from '../find-replace-session';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

/** Find only (Ctrl+F), or Find and Replace (Ctrl+H). */
export type FindReplaceMode = 'find' | 'replace';

export interface FindReplaceControllerState {
	readonly open: boolean;
	readonly mode: FindReplaceMode;
	readonly session: FindReplaceSession;
}

export interface FindReplaceController
	extends PptxController<FindReplaceControllerState>, Omit<FindReplaceSessionActions, 'close'> {
	/** Open the bar in `mode` (searching the kept query again), or switch an open bar's mode. */
	open(mode?: FindReplaceMode): void;
	/** Close the bar: the matches go, the query, replacement and scopes stay. */
	close(): void;
	/**
	 * Ctrl+F / Ctrl+H and the toolbar buttons: close the bar when it is open in `mode`, switch
	 * an open bar to `mode`, or open it in `mode`.
	 */
	toggle(mode?: FindReplaceMode): void;
	/** Open or close from a two-way bound flag (`v-model`, a host prop). */
	setOpen(open: boolean): void;
	/** Route an editor keymap action: `find` and `findReplace` toggle; returns whether it did. */
	handleShortcut(action: string | null | undefined): boolean;
}

export function createFindReplaceController(host: FindReplaceSessionHost): FindReplaceController {
	const store = createControllerStore<FindReplaceControllerState>({
		open: false,
		mode: 'find',
		session: INITIAL_FIND_REPLACE_SESSION,
	});
	const session = findReplaceSessionActions(
		{
			get: () => store.getSnapshot().session,
			set: (next) => store.update((current) => ({ ...current, session: next })),
		},
		host,
	);
	const open = (mode: FindReplaceMode = store.getSnapshot().mode): void => {
		const current = store.getSnapshot();
		if (current.open) {
			if (current.mode !== mode) {
				store.set({ ...current, mode });
			}
			return;
		}
		store.set({ ...current, open: true, mode });
		// The bar shows the query the session kept; search it again from the first match.
		session.search();
	};
	const close = (): void => {
		const current = store.getSnapshot();
		if (current.open) {
			store.set({ ...current, open: false });
		}
		session.close();
	};
	const toggle = (mode: FindReplaceMode = store.getSnapshot().mode): void => {
		const current = store.getSnapshot();
		if (current.open && current.mode === mode) {
			close();
		} else {
			open(mode);
		}
	};
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		...session,
		open,
		close,
		toggle,
		setOpen: (next) => (next ? open() : close()),
		handleShortcut(action) {
			if (action === 'find') {
				toggle('find');
				return true;
			}
			if (action === 'findReplace') {
				toggle('replace');
				return true;
			}
			return false;
		},
	};
}
