import {
	findInSlides,
	findReplacementAvailability,
	findReplaceSessionActions,
	findReplaceStatus,
	INITIAL_FIND_REPLACE_SESSION,
} from 'ooxml-ui/pptx';
import type { FindReplaceOptions, FindReplaceSession, FindReplaceStatus } from 'ooxml-ui/pptx';

import type { Store, ViewerState } from '../state';
import { clampSlideIndex } from '../state';
import type { EditorOps } from './editor-operations';

/**
 * Find & Replace actions for the ribbon's Home > Editing panel. The session
 * (matches, current match, replace plans, the search after a replace) is the
 * shared `find-replace-session` model, kept here; the panel passes its fields
 * on every call and renders {@link FindReplaceActions.status}. A match goes to
 * its slide and selects its element, and every replace is history-integrated
 * like every other editor action.
 *
 * The members added with the session (`next`, `previous`, `status`, `close`)
 * are optional, so a host-built ribbon that implements this interface itself
 * keeps compiling; without them its panel has no Next / Previous.
 */
export interface FindReplaceActions {
	replacementAvailability?(
		query: string,
		matchCase: boolean,
		includeNotes?: boolean,
		includeChartTitles?: boolean,
		includeChartData?: boolean,
	): ReturnType<typeof findReplacementAvailability>;
	/** Search `query` across all slides and go to the first match; returns the match count. */
	search(
		query: string,
		matchCase: boolean,
		includeNotes?: boolean,
		includeChartTitles?: boolean,
		includeChartData?: boolean,
	): number;
	/** Replace the current match of `query`; returns the number replaced (0 or 1). */
	replaceCurrent(
		query: string,
		replacement: string,
		matchCase: boolean,
		includeNotes?: boolean,
		includeChartTitles?: boolean,
		includeChartData?: boolean,
	): number;
	/** Replace every match of `query`; returns the number replaced. */
	replaceAll(
		query: string,
		replacement: string,
		matchCase: boolean,
		includeNotes?: boolean,
		includeChartTitles?: boolean,
		includeChartData?: boolean,
	): number;
	/** Go to the next match, wrapping around. */
	next?(): void;
	/** Go to the previous match, wrapping around. */
	previous?(): void;
	/** The status line's `{{current}} of {{total}}` parameters. */
	status?(): FindReplaceStatus;
	/** The panel closed: drop the matches (the query stays for the next search). */
	close?(): void;
}

export interface FindReplaceActionsDeps {
	store: Store<ViewerState>;
	ops: EditorOps;
}

const scopes = (
	matchCase: boolean,
	includeNotes = false,
	includeChartTitles = true,
	includeChartData = false,
): FindReplaceOptions => ({ matchCase, includeNotes, includeChartTitles, includeChartData });

const sameScopes = (a: FindReplaceOptions, b: FindReplaceOptions): boolean =>
	a.matchCase === b.matchCase &&
	a.includeNotes === b.includeNotes &&
	a.includeChartTitles === b.includeChartTitles &&
	a.includeChartData === b.includeChartData;

export function createFindReplaceActions(deps: FindReplaceActionsDeps): FindReplaceActions {
	const { store, ops } = deps;
	let session: FindReplaceSession = INITIAL_FIND_REPLACE_SESSION;
	const actions = findReplaceSessionActions(
		{
			get: () => session,
			set: (next) => {
				session = next;
			},
		},
		{
			slides: () => store.get().slides,
			commit: (slides) => {
				ops.pushHistory();
				store.set({ slides: [...slides] });
				ops.commitChange();
			},
			navigate: ({ slideIndex, elementId }) => {
				const state = store.get();
				// The editor drops the selection on a slide change, so go first, then select.
				if (state.currentSlide !== slideIndex) {
					store.set({ currentSlide: clampSlideIndex(slideIndex, state.slides.length) });
				}
				ops.select(elementId);
			},
		},
	);
	/** Whether the session already holds the search for this query and these scopes. */
	const holds = (query: string, options: FindReplaceOptions): boolean =>
		session.query === query && sameScopes(session.options, options);
	/** Bring the session to the panel's fields before a replace. */
	const prepare = (query: string, replacement: string, options: FindReplaceOptions): boolean => {
		if (!store.get().editable || !query) {
			return false;
		}
		if (!holds(query, options)) {
			actions.search({ query, options });
		}
		actions.setReplacement(replacement);
		return true;
	};

	return {
		replacementAvailability(query, matchCase, includeNotes, includeChartTitles, includeChartData) {
			const options = scopes(matchCase, includeNotes, includeChartTitles, includeChartData);
			return holds(query, options)
				? findReplacementAvailability(session.matches, session.current)
				: findReplacementAvailability(findInSlides(store.get().slides, query, options), 0);
		},
		search(query, matchCase, includeNotes, includeChartTitles, includeChartData) {
			actions.search({
				query,
				options: scopes(matchCase, includeNotes, includeChartTitles, includeChartData),
			});
			return session.matches.length;
		},
		replaceCurrent(
			query,
			replacement,
			matchCase,
			includeNotes,
			includeChartTitles,
			includeChartData,
		) {
			const options = scopes(matchCase, includeNotes, includeChartTitles, includeChartData);
			return prepare(query, replacement, options) ? actions.replaceCurrent() : 0;
		},
		replaceAll(query, replacement, matchCase, includeNotes, includeChartTitles, includeChartData) {
			const options = scopes(matchCase, includeNotes, includeChartTitles, includeChartData);
			return prepare(query, replacement, options) ? actions.replaceAll() : 0;
		},
		next: actions.next,
		previous: actions.previous,
		status: () => findReplaceStatus(session),
		close: actions.close,
	};
}
