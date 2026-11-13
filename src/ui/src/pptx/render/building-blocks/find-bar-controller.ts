/**
 * The building blocks' Find and Replace bar as a headless controller: whether the host's bar is
 * open plus the shared find session (`find-replace-session`), with the same transitions React's
 * `useFindReplace` makes (a query or scope change searches, closing drops the matches). Ports that
 * keep no find state of their own (vanilla, Vue, Svelte, Angular building blocks) use it as their
 * {@link BuildingBlockFindReplacePort}. DOM-free and framework-free.
 *
 * @module render/building-blocks/find-bar-controller
 */
import { createControllerStore } from '../controller/controller';
import type { PptxController } from '../controller/controller';
import { INITIAL_FIND_REPLACE_SESSION } from '../find-replace-session';
import type { FindReplaceSession } from '../find-replace-session';
import { findReplaceSessionActions } from '../find-replace-session-actions';
import type { FindReplaceSessionHost } from '../find-replace-session-actions';
import type { BuildingBlockFindReplacePort } from './find-layout';

export interface FindBarState {
	readonly open: boolean;
	readonly session: FindReplaceSession;
}

export interface FindBarController extends PptxController<FindBarState> {
	/** The port the session's `actions.findReplace` is built over (the current values). */
	port(): BuildingBlockFindReplacePort;
	setOpen(open: boolean): void;
	toggle(): void;
	/** Tell the bar the host's slides changed (a host whose slides land later, after a replace). */
	slidesChanged(): void;
}

export function createFindBarController(host: FindReplaceSessionHost): FindBarController {
	const store = createControllerStore<FindBarState>({
		open: false,
		session: INITIAL_FIND_REPLACE_SESSION,
	});
	const actions = findReplaceSessionActions(
		{
			get: () => store.getSnapshot().session,
			set: (session) => store.update((current) => ({ ...current, session })),
		},
		host,
	);
	const setOpen = (open: boolean): void => {
		if (store.getSnapshot().open === open) {
			return;
		}
		// Closing drops the matches, which also drive the canvas highlight.
		if (!open) {
			actions.close();
		}
		store.update((current) => ({ ...current, open }));
	};
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		setOpen,
		toggle: () => setOpen(!store.getSnapshot().open),
		slidesChanged: () => actions.slidesChanged(),
		port() {
			const { open, session } = store.getSnapshot();
			return {
				open,
				setOpen,
				query: session.query,
				replacement: session.replacement,
				options: session.options,
				matches: session.matches,
				currentIndex: session.current,
				setQuery: (query) => actions.search({ query }),
				setReplacement: (replacement) => actions.setReplacement(replacement),
				setOptions: (change) => actions.search({ options: change }),
				find: () => actions.search(),
				navigate: (delta) => (delta > 0 ? actions.next() : actions.previous()),
				replaceCurrent: () => void actions.replaceCurrent(),
				replaceAll: () => void actions.replaceAll(),
			};
		},
	};
}
