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
/**
 * Visio opens Find from Home > Editing rather than keeping a box in the ribbon. The find bar
 * sits under the ribbon, hidden until Find... or Ctrl+F, and closes with its button or Escape.
 */
export function createFindBar(doc: Document): HTMLElement {
	const bar = doc.createElement('div');
	bar.className = 'find-bar';
	bar.hidden = true;
	const close = doc.createElement('button');
	close.type = 'button';
	close.className = 'find-close';
	close.dataset.action = 'find-close';
	close.setAttribute('aria-label', 'Close Find');
	close.title = 'Close Find (Esc)';
	close.textContent = '×';
	bar.append(createSearchGroup(doc), close);
	return bar;
}
/** The find bar closes from its button, or Escape once the query is already empty. */
export function wireFindBar(
	bar: HTMLElement,
	input: HTMLInputElement,
	closed: () => void,
): () => void {
	const Abort = bar.ownerDocument.defaultView?.AbortController ?? AbortController;
	const events = new Abort();
	const close = () => {
		bar.hidden = true;
		closed();
	};
	bar.addEventListener(
		'click',
		(event) => {
			if ((event.target as Element).closest?.('[data-action="find-close"]')) close();
		},
		{ signal: events.signal },
	);
	// Capture runs before the search field clears its query on Escape.
	bar.addEventListener(
		'keydown',
		(event) => {
			if (event.key === 'Escape' && input.value === '') close();
		},
		{ signal: events.signal, capture: true },
	);
	return () => events.abort();
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
