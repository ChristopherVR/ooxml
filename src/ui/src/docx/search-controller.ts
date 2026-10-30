import { closeHistory } from 'prosemirror-history';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	expectDefined,
	FindPatternError,
	replacementParts,
	type FindOptions,
} from 'ooxml-core/docx/ui';
import {
	findDocumentMatches,
	replaceTargets,
	type DocumentMatch,
	type ReplacementTarget,
} from './search-document';

export interface SearchMatch {
	from: number;
	to: number;
}

export interface SearchStatus {
	query: string;
	count: number;
	active: number;
	readOnly: boolean;
	/** The query is a wildcard pattern Word would reject; nothing matches it. */
	invalid: boolean;
}

/** The Find options of the dialog: Match case, whole words, wildcards and the rest of More >>. */
export type SearchOptions = FindOptions;

export class SearchController {
	private query = '';
	private options: SearchOptions = {};
	private matches: DocumentMatch[] = [];
	private activeIndex = -1;
	private invalid = false;

	constructor(private readonly getView: () => EditorView | undefined) {}

	get status(): SearchStatus {
		return {
			query: this.query,
			count: this.matches.length,
			active: this.activeIndex < 0 ? 0 : this.activeIndex + 1,
			readOnly: !this.getView()?.editable,
			invalid: this.invalid,
		};
	}

	search(query: string, options: SearchOptions = this.options): SearchStatus {
		this.query = query;
		this.options = options;
		return this.refresh();
	}

	refresh(): SearchStatus {
		const view = this.getView();
		this.matches = [];
		this.invalid = false;
		if (view) {
			try {
				this.matches = findDocumentMatches(view.state.doc, this.query, this.options);
			} catch (error) {
				if (!(error instanceof FindPatternError)) throw error;
				this.invalid = true;
			}
		}
		this.activeIndex = this.indexAtSelection();
		return this.status;
	}

	findNext(): SearchMatch | null {
		const view = this.getView();
		this.refresh();
		if (!view || !this.matches.length) return null;
		let next = this.activeIndex + 1;
		if (this.activeIndex < 0) {
			const cursor = view.state.selection.from;
			next = this.matches.findIndex((match) => match.from >= cursor);
			if (next < 0) next = 0;
		}
		return this.select(next % this.matches.length);
	}

	findPrevious(): SearchMatch | null {
		const view = this.getView();
		this.refresh();
		if (!view || !this.matches.length) return null;
		let previous = this.activeIndex - 1;
		if (this.activeIndex < 0) {
			const cursor = view.state.selection.from;
			previous = -1;
			for (let index = this.matches.length - 1; index >= 0; index--) {
				const match = this.matches[index];
				if (match && match.from < cursor) {
					previous = index;
					break;
				}
			}
			if (previous < 0) previous = this.matches.length - 1;
		}
		return this.select((previous + this.matches.length) % this.matches.length);
	}

	replaceOne(replacement: string): boolean {
		const view = this.getView();
		if (!view?.editable || !this.query) return false;
		this.refresh();
		if (this.activeIndex < 0 && !this.findNext()) return false;
		const match = this.matches[this.activeIndex];
		if (!match) return false;
		return this.replace(view, [match], replacement);
	}

	replaceAll(replacement: string): number {
		const view = this.getView();
		if (!view?.editable || !this.query) return 0;
		this.refresh();
		if (!this.matches.length) return 0;
		const targets = [...this.matches];
		return this.replace(view, targets, replacement) ? targets.length : 0;
	}

	private indexAtSelection(): number {
		const selection = this.getView()?.state.selection;
		if (!selection) return -1;
		return this.matches.findIndex(
			(match) => match.from === selection.from && match.to === selection.to,
		);
	}

	private select(index: number): SearchMatch {
		const view = this.getView()!;
		const match = expectDefined(this.matches[index], 'search match');
		this.activeIndex = index;
		view.dispatch(
			view.state.tr
				.setSelection(TextSelection.create(view.state.doc, match.from, match.to))
				.scrollIntoView(),
		);
		return match;
	}

	/** Replaces the matches (in document order) in one undoable step; false if the replacement is invalid. */
	private replace(view: EditorView, matches: DocumentMatch[], replacement: string): boolean {
		const wildcards = Boolean(this.options.wildcards);
		let targets: ReplacementTarget[];
		try {
			targets = matches.map((match) => ({
				from: match.from,
				to: match.to,
				parts: replacementParts(replacement, match, wildcards),
			}));
		} catch (error) {
			if (!(error instanceof FindPatternError)) throw error;
			this.invalid = true;
			return false;
		}
		const transaction = replaceTargets(view.state, targets);
		// The cursor goes after the first replacement, mapped through every step of the transaction.
		const cursor = transaction.mapping.map(expectDefined(matches[0], 'first match').to, 1);
		const selection = TextSelection.near(transaction.doc.resolve(cursor));
		view.dispatch(closeHistory(transaction.setSelection(selection).scrollIntoView()));
		this.refresh();
		return true;
	}
}

/** Every match of a query in the document, as ranges (Find without its replacement details). */
export function findTextMatches(
	doc: Parameters<typeof findDocumentMatches>[0],
	query: string,
	options: SearchOptions = {},
): SearchMatch[] {
	return findDocumentMatches(doc, query, options).map(({ from, to }) => ({ from, to }));
}
