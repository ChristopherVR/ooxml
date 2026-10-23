/**
 * Shared session model for the PowerPoint Find and Replace bar.
 *
 * Searching and replacing are the core's (`findInSlides`, `applyFindReplacements`). What the bar
 * does with them while it is open is the same in every binding, so it lives here: the query, the
 * replacement and the four scopes, the match list and the current match, Next and Previous with
 * wrap-around, the "n of N" status, Replace and Replace All, the search that runs again once the
 * replaced slides land (and not on every ordinary edit), and the matches a closed bar drops.
 *
 * The session is a plain value each host keeps in its own reactive container:
 * {@link reduceFindReplaceSession} is the only place it changes; {@link findReplaceStatus},
 * {@link findReplaceTarget} and {@link planFindReplace} derive from it; and
 * `findReplaceSessionActions` (`find-replace-session-actions.ts`) binds the bar's actions to the
 * host's store, its slides, its commit and its navigation.
 *
 * DOM-free and framework-free.
 */
import type { PptxSlide } from 'ooxml-core/pptx';

import { findInSlides } from './find-replace';
import type { FindResult } from './find-replace';

/** The bar's four search scopes. */
export interface FindReplaceOptions {
	readonly matchCase: boolean;
	readonly includeNotes: boolean;
	readonly includeChartTitles: boolean;
	readonly includeChartData: boolean;
}

/** The scopes a new bar starts with, in every binding. */
export const DEFAULT_FIND_REPLACE_OPTIONS: FindReplaceOptions = Object.freeze({
	matchCase: false,
	includeNotes: false,
	includeChartTitles: true,
	includeChartData: false,
});

/** Replace (the current match) or Replace All. */
export type FindReplaceScope = 'current' | 'all';

/** A replace was applied; the search runs again once the host's slides are no longer `base`. */
export interface FindReplacePendingRefind {
	readonly scope: FindReplaceScope;
	/** The plan's own search of `base`: where the replaced match was in it, and how many it found. */
	readonly index: number;
	readonly count: number;
	/** The slides the replace was applied to. */
	readonly base: readonly PptxSlide[];
}

export interface FindReplaceSession {
	readonly query: string;
	readonly replacement: string;
	readonly options: FindReplaceOptions;
	readonly matches: readonly FindResult[];
	/** Zero-based current match; -1 when there is none. */
	readonly current: number;
	readonly pendingRefind: FindReplacePendingRefind | null;
}

/** A session before the user has typed anything. */
export const INITIAL_FIND_REPLACE_SESSION: FindReplaceSession = Object.freeze({
	query: '',
	replacement: '',
	options: DEFAULT_FIND_REPLACE_OPTIONS,
	matches: Object.freeze([]),
	current: -1,
	pendingRefind: null,
});

/** A change to the query or the scopes; whatever it leaves out stays as it is. */
export interface FindReplaceSearchChange {
	readonly query?: string;
	readonly options?: Partial<FindReplaceOptions>;
}

/** Which matches to apply, found again in the slides the plan was made for. */
export interface FindReplacePlan {
	readonly scope: FindReplaceScope;
	readonly matches: readonly FindResult[];
	readonly replacement: string;
	/**
	 * That search: how many matches it found, and the current match's place among them (-1 for
	 * Replace All). The session's own list can be stale (an ordinary edit leaves it alone), so the
	 * search after the replace counts from these.
	 */
	readonly count: number;
	readonly index: number;
}

export type FindReplaceSessionEvent =
	/** Search `slides` (after applying the change); the cursor goes to the first match. */
	| ({ readonly type: 'search'; readonly slides: readonly PptxSlide[] } & FindReplaceSearchChange)
	| { readonly type: 'setReplacement'; readonly replacement: string }
	/** Next (1) or Previous (-1), wrapping around the ends. */
	| { readonly type: 'step'; readonly direction: 1 | -1 }
	/** `plan` was applied to `slides`: search again once the host's slides have changed. */
	| {
			readonly type: 'replaced';
			readonly plan: FindReplacePlan;
			readonly slides: readonly PptxSlide[];
	  }
	/** The host's slides are now `slides`. Searches again only while a replace waits for them. */
	| { readonly type: 'slidesChanged'; readonly slides: readonly PptxSlide[] }
	/** The bar closed: the matches go; the query, replacement and scopes stay for the next open. */
	| { readonly type: 'close' };

function search(slides: readonly PptxSlide[], session: FindReplaceSession): FindResult[] {
	return session.query ? findInSlides(slides, session.query, session.options) : [];
}

/** Where the cursor lands once a replace has been searched again. */
function cursorAfterReplace(pending: FindReplacePendingRefind, count: number): number {
	if (count === 0) {
		return -1;
	}
	if (pending.scope === 'all') {
		return 0;
	}
	// Only the replaced match went, so the one after it now has its index. Matches the replacement
	// text itself created sit at that index too: step past them, as PowerPoint moves on to the next
	// occurrence instead of matching what it just wrote.
	const created = Math.max(0, count - (pending.count - 1));
	return (pending.index + created) % count;
}

/** The one transition function. Returns `session` itself when the event changes nothing. */
export function reduceFindReplaceSession(
	session: FindReplaceSession,
	event: FindReplaceSessionEvent,
): FindReplaceSession {
	switch (event.type) {
		case 'search': {
			const next: FindReplaceSession = {
				...session,
				query: event.query ?? session.query,
				options: event.options ? { ...session.options, ...event.options } : session.options,
				pendingRefind: null,
			};
			const matches = search(event.slides, next);
			return { ...next, matches, current: matches.length > 0 ? 0 : -1 };
		}
		case 'setReplacement':
			return event.replacement === session.replacement
				? session
				: { ...session, replacement: event.replacement };
		case 'step': {
			const count = session.matches.length;
			if (count === 0) {
				return session;
			}
			const from = session.current < 0 ? (event.direction > 0 ? -1 : 0) : session.current;
			const current = (from + event.direction + count) % count;
			return current === session.current ? session : { ...session, current };
		}
		case 'replaced':
			return {
				...session,
				pendingRefind: {
					scope: event.plan.scope,
					index: event.plan.index,
					count: event.plan.count,
					base: event.slides,
				},
			};
		case 'slidesChanged': {
			const pending = session.pendingRefind;
			if (!pending || event.slides === pending.base) {
				return session;
			}
			const matches = search(event.slides, session);
			return {
				...session,
				matches,
				current: cursorAfterReplace(pending, matches.length),
				pendingRefind: null,
			};
		}
		case 'close':
			return session.matches.length === 0 && session.current === -1 && !session.pendingRefind
				? session
				: { ...session, matches: [], current: -1, pendingRefind: null };
	}
}

/** The `{{current}} of {{total}}` parameters: `current` is 1-based, 0 when there is no match. */
export interface FindReplaceStatus {
	readonly current: number;
	readonly total: number;
}

export function findReplaceStatus(
	session: Pick<FindReplaceSession, 'matches' | 'current'>,
): FindReplaceStatus {
	const total = session.matches.length;
	return { current: total > 0 && session.current >= 0 ? session.current + 1 : 0, total };
}

/** Where a match is: the host goes to the slide and selects the element (none for notes). */
export interface FindReplaceTarget {
	readonly slideIndex: number;
	readonly elementId: string | null;
}

/** The current match as a navigation descriptor, or null when there is none. */
export function findReplaceTarget(
	session: Pick<FindReplaceSession, 'matches' | 'current'>,
): FindReplaceTarget | null {
	const match = session.current >= 0 ? session.matches[session.current] : undefined;
	if (!match) {
		return null;
	}
	return { slideIndex: match.slideIndex, elementId: match.notes ? null : match.elementId };
}

/** Everything that tells two matches apart: the text object and the position in it. */
function matchKey(match: FindResult): string {
	return JSON.stringify([
		match.slideIndex,
		match.elementId,
		match.elementPath ?? null,
		match.tableCell ?? null,
		match.smartArtNodeId ?? null,
		match.notes ?? false,
		match.chartTitle ?? false,
		match.chartSeriesIndex ?? null,
		match.chartCategoryIndex ?? null,
		match.segmentIndex,
		match.startOffset,
		match.length,
	]);
}

/**
 * What Replace (`current`) or Replace All (`all`) would apply to `slides`, or null when it is
 * refused: there is nothing to replace, a match is read-only (linked or shared chart text), or the
 * current match is no longer in `slides` (the deck changed after the search, so its offsets would
 * no longer point at the query).
 */
export function planFindReplace(
	session: FindReplaceSession,
	slides: readonly PptxSlide[],
	scope: FindReplaceScope,
): FindReplacePlan | null {
	const found = search(slides, session);
	let matches: readonly FindResult[] = found;
	let index = -1;
	if (scope === 'current') {
		const current = session.current >= 0 ? session.matches[session.current] : undefined;
		const key = current ? matchKey(current) : null;
		index = key === null ? -1 : found.findIndex((match) => matchKey(match) === key);
		matches = index >= 0 ? [found[index]!] : [];
	}
	if (matches.length === 0 || matches.some((match) => match.readOnly)) {
		return null;
	}
	return { scope, matches, replacement: session.replacement, count: found.length, index };
}
