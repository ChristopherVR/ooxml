/**
 * The building blocks' Find and Replace bar: the shared `createFindReplaceController` (the bar's
 * open flag and mode plus the find session) with a {@link BuildingBlockFindReplacePort} on top.
 * Ports that keep no find state of their own (vanilla, Vue, Svelte, Angular building blocks) use
 * it as their port. DOM-free and framework-free.
 *
 * @module render/building-blocks/find-bar-controller
 */
import { createFindReplaceController } from '../controller/find-replace-controller';
import type { FindReplaceController } from '../controller/find-replace-controller';
import type { FindReplaceSessionHost } from '../find-replace-session-actions';
import type { BuildingBlockFindReplacePort } from './find-layout';

/** The bar's state: the shared find and replace controller's. */
export type FindBarState = ReturnType<FindReplaceController['getSnapshot']>;

export interface FindBarController extends FindReplaceController {
	/** The port the session's `actions.findReplace` is built over (the current values). */
	port(): BuildingBlockFindReplacePort;
}

/** The shared `createFindReplaceController` with the building blocks' port on top. */
export function createFindBarController(host: FindReplaceSessionHost): FindBarController {
	const bar = createFindReplaceController(host);
	return {
		...bar,
		port() {
			const { open, session } = bar.getSnapshot();
			return {
				open,
				setOpen: bar.setOpen,
				query: session.query,
				replacement: session.replacement,
				options: session.options,
				matches: session.matches,
				currentIndex: session.current,
				setQuery: (query) => bar.search({ query }),
				setReplacement: (replacement) => bar.setReplacement(replacement),
				setOptions: (change) => bar.search({ options: change }),
				find: () => bar.search(),
				navigate: (delta) => (delta > 0 ? bar.next() : bar.previous()),
				replaceCurrent: () => void bar.replaceCurrent(),
				replaceAll: () => void bar.replaceAll(),
			};
		},
	};
}
