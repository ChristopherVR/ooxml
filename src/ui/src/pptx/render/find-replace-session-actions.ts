/**
 * The Find and Replace bar's actions over a host's store, for {@link FindReplaceSession}.
 *
 * Each action is a transition of `reduceFindReplaceSession` plus what the host has to do about it:
 * commit replaced slides, and show the current match. DOM-free and framework-free.
 */
import type { PptxSlide } from 'ooxml-core/pptx';

import { applyFindReplacements } from './find-replace';
import {
	findReplaceTarget,
	planFindReplace,
	reduceFindReplaceSession,
} from './find-replace-session';
import type {
	FindReplaceScope,
	FindReplaceSearchChange,
	FindReplaceSession,
	FindReplaceSessionEvent,
	FindReplaceTarget,
} from './find-replace-session';

/** Where a host keeps its session. `get` must return what the last `set` stored. */
export interface FindReplaceSessionStore {
	get(): FindReplaceSession;
	set(next: FindReplaceSession): void;
}

/** What the actions read from and do to the host. */
export interface FindReplaceSessionHost {
	/** The slides as they are now. */
	slides(): readonly PptxSlide[];
	/** Commit replaced slides as one undoable edit (history, dirty flag). */
	commit(slides: readonly PptxSlide[], scope: FindReplaceScope): void;
	/** Show a match: go to its slide and select its element (or nothing, for `null`). */
	navigate(target: FindReplaceTarget): void;
}

/** A host's bar actions, each one a transition plus what the host has to do about it. */
export interface FindReplaceSessionActions {
	/** Search with the change applied (a new query, a toggled scope) or, with none, again. */
	search(change?: FindReplaceSearchChange): void;
	setReplacement(replacement: string): void;
	next(): void;
	previous(): void;
	/** Replace the current match; returns how many replacements landed (0 when refused). */
	replaceCurrent(): number;
	replaceAll(): number;
	/** Tell the session the host's slides changed; it searches again only after a replace. */
	slidesChanged(): void;
	close(): void;
}

/**
 * The bar's actions over a host's store. A search, Next, Previous and the search after a replace
 * show the current match through `host.navigate`. After a commit the actions read `host.slides()`
 * again: a host whose state updates synchronously is searched again at once, and one that updates
 * later (React) calls {@link FindReplaceSessionActions.slidesChanged} when its slides land.
 */
export function findReplaceSessionActions(
	store: FindReplaceSessionStore,
	host: FindReplaceSessionHost,
): FindReplaceSessionActions {
	/** Apply one event; true when it changed the session. */
	const apply = (event: FindReplaceSessionEvent): boolean => {
		const current = store.get();
		const next = reduceFindReplaceSession(current, event);
		if (next !== current) {
			store.set(next);
		}
		return next !== current;
	};
	const show = (): void => {
		const target = findReplaceTarget(store.get());
		if (target) {
			host.navigate(target);
		}
	};
	const step = (direction: 1 | -1): void => {
		// A lone match stays current, and is shown again.
		apply({ type: 'step', direction });
		show();
	};
	const slidesChanged = (): void => {
		if (apply({ type: 'slidesChanged', slides: host.slides() })) {
			show();
		}
	};
	const replace = (scope: FindReplaceScope): number => {
		const slides = host.slides();
		const plan = planFindReplace(store.get(), slides, scope);
		if (!plan) {
			return 0;
		}
		const result = applyFindReplacements(slides, plan.matches, plan.replacement);
		if (result.replacements === 0) {
			return 0;
		}
		host.commit(result.slides, scope);
		apply({ type: 'replaced', plan, slides });
		slidesChanged();
		return result.replacements;
	};
	return {
		search: (change = {}) => {
			apply({ type: 'search', slides: host.slides(), ...change });
			show();
		},
		setReplacement: (replacement) => {
			apply({ type: 'setReplacement', replacement });
		},
		next: () => step(1),
		previous: () => step(-1),
		replaceCurrent: () => replace('current'),
		replaceAll: () => replace('all'),
		slidesChanged,
		close: () => {
			apply({ type: 'close' });
		},
	};
}
