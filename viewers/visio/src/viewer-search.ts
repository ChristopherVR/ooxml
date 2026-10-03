import { TEXT_SEARCH_LIMITS, textSearchStatus } from './document-text-search.js';
import type { ViewerState } from './controller.js';

/** Builds the Find group: a search field, previous/next result buttons and a live status. */
export function createSearchGroup(doc: Document): HTMLElement {
	const group = doc.createElement('div');
	group.className = 'ribbon-group search-group';
	group.setAttribute('role', 'group');
	group.setAttribute('aria-label', 'Find');
	const controls = doc.createElement('div');
	controls.className = 'search-controls';
	controls.setAttribute('role', 'search');
	controls.setAttribute('aria-label', 'Diagram text search');
	const label = doc.createElement('label');
	label.append('Find ');
	const input = doc.createElement('input');
	input.type = 'search';
	input.setAttribute('aria-label', 'Search diagram text');
	input.placeholder = 'Text across pages';
	input.maxLength = TEXT_SEARCH_LIMITS.queryCharacters;
	input.autocomplete = 'off';
	input.spellcheck = false;
	input.setAttribute('aria-describedby', 'search-status');
	label.append(input);
	const step = (action: string, name: string, text: string) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.dataset.action = action;
		button.setAttribute('aria-label', name);
		button.textContent = text;
		return button;
	};
	const status = doc.createElement('span');
	status.id = 'search-status';
	status.setAttribute('role', 'status');
	status.setAttribute('aria-live', 'polite');
	status.setAttribute('aria-atomic', 'true');
	controls.append(
		label,
		step('search-previous', 'Previous matching shape', 'Previous'),
		step('search-next', 'Next matching shape', 'Next'),
		status,
	);
	const caption = doc.createElement('span');
	caption.className = 'group-label';
	caption.textContent = 'Find in drawing';
	group.append(controls, caption);
	return group;
}
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
