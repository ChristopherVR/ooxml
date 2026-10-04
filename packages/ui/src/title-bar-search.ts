import { createIconSvg, paintIcon } from './icons.js';

/** One command the title-bar search can run. `category` shows on the right of its row. */
export interface OfficeTitleBarCommand {
	id: string;
	label: string;
	category?: string | undefined;
}

/** The centred command search. Every string arrives translated from the host. */
export interface OfficeTitleBarSearch {
	placeholder: string;
	label: string;
	/** Heading over the matches and the live-region text when there are some. */
	heading: string;
	/** Row and live-region text when nothing matches. */
	empty: string;
	/** Searched by label when `match` is absent. */
	commands?: readonly OfficeTitleBarCommand[] | undefined;
	/** Host ranking (synonyms, keywords); replaces the default label search. */
	match?: ((query: string) => readonly OfficeTitleBarCommand[]) | undefined;
	/** Text of the trailing "search the document" row; absent drops that row. */
	content?: ((query: string) => string) | undefined;
	/** Glyph of the content row; defaults to `search`. */
	contentIcon?: string | undefined;
}

/** Detail of the search request: a chosen command, or a document search when `command` is absent. */
export interface OfficeTitleBarSearchDetail {
	query: string;
	command?: string;
}

/** The dropdown shows at most this many commands so it never outgrows its box. */
export const OFFICE_TITLE_BAR_SEARCH_LIMIT = 8;

function labelMatches(search: OfficeTitleBarSearch, query: string): OfficeTitleBarCommand[] {
	const needle = query.trim().toLowerCase();
	return (search.commands ?? []).filter((entry) => entry.label.toLowerCase().includes(needle));
}

export interface TitleBarSearchView {
	box: HTMLElement;
	input: HTMLElement & { value: string };
	render(search: OfficeTitleBarSearch | undefined): void;
	close(): void;
}

/**
 * A search field plus its result list. The query stays local to the view, so hosts never
 * round-trip keystrokes; they receive one request per commit.
 */
export function createTitleBarSearch(
	doc: Document,
	fieldTag: string,
	request: (detail: OfficeTitleBarSearchDetail) => void,
): TitleBarSearchView {
	let model: OfficeTitleBarSearch | undefined;
	const box = doc.createElement('div');
	box.className = 'box';
	const input = doc.createElement(fieldTag) as HTMLElement & { value: string };
	input.setAttribute('variant', 'titlebar');
	input.setAttribute('part', 'search');
	const results = doc.createElement('div');
	results.className = 'results';
	results.setAttribute('role', 'listbox');
	results.hidden = true;
	const live = doc.createElement('div');
	live.className = 'sr';
	live.setAttribute('role', 'status');
	box.append(input, results, live);

	let matches: OfficeTitleBarCommand[] = [];
	let active = 0;
	const close = (): void => {
		results.hidden = true;
		results.replaceChildren();
		live.textContent = '';
		matches = [];
	};
	const commit = (entry?: OfficeTitleBarCommand): void => {
		const query = input.value;
		input.value = '';
		close();
		request(entry ? { query, command: entry.id } : { query });
	};
	const row = (className: string, text: string): HTMLElement => {
		const node = doc.createElement('div');
		node.className = className;
		node.textContent = text;
		return node;
	};
	const option = (className: string, onPick: () => void): HTMLButtonElement => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.tabIndex = -1;
		button.className = className;
		// mousedown (not click) so the choice lands before the field blurs.
		button.addEventListener('mousedown', (event) => {
			event.preventDefault();
			onPick();
		});
		return button;
	};
	const paint = (): void => {
		[...results.querySelectorAll('[role="option"]')].forEach((node, index) =>
			node.setAttribute('aria-selected', String(index === active)),
		);
	};
	const open = (): void => {
		const query = input.value;
		if (!model || !query.trim()) {
			close();
			return;
		}
		const found = model.match ? [...model.match(query)] : labelMatches(model, query);
		matches = found.slice(0, OFFICE_TITLE_BAR_SEARCH_LIMIT);
		active = Math.min(active, Math.max(matches.length - 1, 0));
		results.replaceChildren();
		if (matches.length > 0) {
			results.append(row('heading', model.heading));
			matches.forEach((entry, index) => {
				const button = option('', () => commit(entry));
				button.setAttribute('role', 'option');
				const label = doc.createElement('span');
				label.textContent = entry.label;
				const category = doc.createElement('span');
				category.className = 'cat';
				category.textContent = entry.category ?? '';
				button.append(label, category);
				button.addEventListener('mouseenter', () => {
					active = index;
					paint();
				});
				results.append(button);
			});
		} else {
			results.append(row('empty', model.empty));
		}
		if (model.content) {
			const content = option('content', () => commit());
			const svg = createIconSvg(doc);
			paintIcon(svg, model.contentIcon ?? 'search');
			const text = doc.createElement('span');
			text.textContent = model.content(query);
			content.append(svg, text);
			results.append(content);
		}
		results.hidden = false;
		live.textContent = matches.length > 0 ? model.heading : model.empty;
		paint();
	};

	input.addEventListener('input', () => {
		active = 0;
		open();
	});
	input.addEventListener('focusin', open);
	box.addEventListener('focusout', (event) => {
		const next = event.relatedTarget as Node | null;
		if (!next || !box.contains(next)) close();
	});
	input.addEventListener('keydown', (event) => {
		const { key, ctrlKey, metaKey } = event as KeyboardEvent;
		if (key === 'Enter' && input.value.trim()) {
			commit(matches[active]);
		} else if (key === 'Escape' && input.value) {
			input.value = '';
			close();
		} else if ((key === 'ArrowDown' || key === 'ArrowUp') && matches.length > 0) {
			event.preventDefault();
			active = (active + (key === 'ArrowDown' ? 1 : matches.length - 1)) % matches.length;
			paint();
		}
		// Typing must not reach the host document's shortcuts.
		if (!ctrlKey && !metaKey) event.stopPropagation();
	});

	return {
		box,
		input,
		render(search) {
			model = search;
			box.hidden = !search;
			if (!search) {
				close();
				return;
			}
			input.setAttribute('placeholder', search.placeholder);
			input.setAttribute('aria-label', search.label);
			if (!results.hidden) open();
		},
		close,
	};
}
