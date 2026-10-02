import { TEXT_SEARCH_LIMITS, textSearchStatus } from './document-text-search.js';
import type { ViewerState } from './controller.js';

export const searchTemplate = `<div class="search-controls" role="search" aria-label="Diagram text search"><label>Find <input type="search" aria-label="Search diagram text" placeholder="Text across pages" maxlength="${TEXT_SEARCH_LIMITS.queryCharacters}" autocomplete="off" spellcheck="false" aria-describedby="search-status"></label><button type="button" data-action="search-previous" aria-label="Previous matching shape">Previous</button><button type="button" data-action="search-next" aria-label="Next matching shape">Next</button><span id="search-status" role="status" aria-live="polite" aria-atomic="true"></span></div>`;
export interface SearchControls {
	input: HTMLInputElement;
	status: HTMLSpanElement;
	previous: HTMLButtonElement;
	next: HTMLButtonElement;
}
export function searchControls(root: ShadowRoot): SearchControls {
	return {
		input: root.querySelector('input[type="search"]')!,
		status: root.querySelector('#search-status')!,
		previous: root.querySelector('[data-action="search-previous"]')!,
		next: root.querySelector('[data-action="search-next"]')!,
	};
}
export function renderSearchControls(controls: SearchControls, state: ViewerState): void {
	if (controls.input.value !== state.search.query) controls.input.value = state.search.query;
	controls.input.disabled = !state.document?.pages.length;
	controls.previous.disabled = controls.next.disabled = !state.search.results.length;
	controls.status.textContent = state.document
		? textSearchStatus(state.search)
		: 'Open a diagram to search';
	const result = state.search.results[state.search.activeIndex];
	// The result preview and document labels remain inert text, including HTML-like strings.
	controls.status.title = result
		? `${result.pageName} · ${result.shapeName}: ${result.preview}`
		: '';
}
