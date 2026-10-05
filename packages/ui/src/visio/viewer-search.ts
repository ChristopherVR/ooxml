import { TEXT_SEARCH_LIMITS, textSearchStatus } from './document-text-search.js';
import type { ViewerController, ViewerState } from './controller.js';

export type FindBar = HTMLElement & {
	open: boolean;
	value: string;
	status: string;
	statusTitle: string;
	show(): void;
	close(): void;
};

/**
 * Visio opens Find from Home > Editing rather than keeping a box in the ribbon. The shared
 * `office-ui-find-bar` sits under the ribbon, hidden until Find... or Ctrl+F; Enter and
 * Shift+Enter step through matching shapes and Escape clears, then closes.
 */
export function createFindBar(doc: Document): FindBar {
	const bar = doc.createElement('office-ui-find-bar') as FindBar;
	bar.className = 'find-bar';
	bar.setAttribute('label', 'Diagram text search');
	bar.setAttribute('input-label', 'Search diagram text');
	bar.setAttribute('placeholder', 'Text across pages');
	bar.setAttribute('maxlength', String(TEXT_SEARCH_LIMITS.queryCharacters));
	bar.setAttribute('previous-label', 'Previous matching shape');
	bar.setAttribute('next-label', 'Next matching shape');
	return bar;
}

/** Route the find bar's queries and steps to the controller; `closed` restores focus. */
export function wireFindBar(
	bar: FindBar,
	controller: ViewerController,
	closed: () => void,
): () => void {
	const Abort = bar.ownerDocument.defaultView?.AbortController ?? AbortController;
	const events = new Abort();
	const options = { signal: events.signal };
	bar.addEventListener(
		'office-find-input',
		(event) => controller.setSearchQuery((event as CustomEvent<{ query: string }>).detail.query),
		options,
	);
	bar.addEventListener(
		'office-find-step',
		(event) => {
			const { direction } = (event as CustomEvent<{ direction: 'next' | 'previous' }>).detail;
			if (direction === 'next') controller.nextSearchResult();
			else controller.previousSearchResult();
		},
		options,
	);
	bar.addEventListener('office-find-close', closed, options);
	return () => events.abort();
}

export function renderFindBar(bar: FindBar, state: ViewerState): void {
	bar.value = state.search.query;
	bar.toggleAttribute('disabled', !state.document?.pages.length);
	bar.toggleAttribute('navigation-disabled', !state.search.results.length);
	bar.status = state.document ? textSearchStatus(state.search) : 'Open a diagram to search';
	const result = state.search.results[state.search.activeIndex];
	// The result preview and document labels remain inert text, including HTML-like strings.
	bar.statusTitle = result ? `${result.pageName} · ${result.shapeName}: ${result.preview}` : '';
}
