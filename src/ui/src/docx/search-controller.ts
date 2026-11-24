import { NodeSelection, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { hasFormatCriteria, type ReplaceFormatPatch } from 'ooxml-core/docx';
import { expectDefined, FindPatternError, type FindOptions } from 'ooxml-core/docx/ui';
import { findDocumentMatches } from './search-document';
import {
	findUnitMatches,
	mainUnits,
	searchUnits,
	type SearchUnit,
	type StoryMatch,
} from './search-stories';
import { applyReplacements, type SearchStoryHost } from './search-story-replace';

export type { SearchStoryHost } from './search-story-replace';

export interface SearchMatch {
	from: number;
	to: number;
}

/** The story a match is in: what the Find dialog names beside its count. */
export interface SearchStoryInfo {
	readonly kind: 'main' | 'footnote' | 'endnote' | 'textBoxes' | 'header' | 'footer';
	/** Headers and footers: `default`, `first` or `even`. */
	readonly slot?: 'default' | 'first' | 'even';
	/** Headers and footers: the zero-based section that owns the part. */
	readonly sectionIndex?: number;
}

export interface SearchStatus {
	query: string;
	/** True when active Format > Font/Paragraph/Style criteria are set (enables format-only Find). */
	hasFormat: boolean;
	count: number;
	active: number;
	readOnly: boolean;
	/** The query is a wildcard pattern Word would reject; nothing matches it. */
	invalid: boolean;
	/** The story of the active match, when there is one. */
	story?: SearchStoryInfo;
}

/** The Find options of the dialog: Match case, whole words, wildcards and the rest of More >>. */
export type SearchOptions = FindOptions;

/** A match Find selected outside an editor caret (a text box, a story it could not open). */
interface Placed {
	readonly key: string;
	readonly from: number;
	readonly to: number;
	/** The body selection right after placing it; any other selection means the user moved on. */
	readonly body: SearchMatch;
}

/**
 * Find and Replace over the body and, given a story host, the other stories in Word's order (see
 * `findStories`): Find Next runs off the end of one story into the next and wraps to the body.
 */
export class SearchController {
	private query = '';
	private options: SearchOptions = {};
	private units: SearchUnit[] = [];
	private matches: StoryMatch[] = [];
	private activeIndex = -1;
	private invalid = false;
	private placed: Placed | undefined;

	constructor(
		private readonly getView: () => EditorView | undefined,
		private readonly stories?: SearchStoryHost,
	) {}

	get status(): SearchStatus {
		const story = this.matches[this.activeIndex]?.unit.story;
		return {
			query: this.query,
			hasFormat: hasFormatCriteria(this.options.format),
			count: this.matches.length,
			active: this.activeIndex < 0 ? 0 : this.activeIndex + 1,
			readOnly: !this.getView()?.editable,
			invalid: this.invalid,
			...(story && {
				story:
					story.kind === 'header' || story.kind === 'footer'
						? { kind: story.kind, slot: story.slot, sectionIndex: story.sectionIndex }
						: { kind: story.kind },
			}),
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
		this.units = [];
		this.invalid = false;
		if (view && (this.query || hasFormatCriteria(this.options.format))) {
			const open = this.stories?.openStory();
			const units = (this.units = this.stories
				? searchUnits(
						view.state.doc,
						this.stories.model(),
						open && { key: open.key, doc: open.view.state.doc },
					)
				: mainUnits(view.state.doc));
			try {
				this.matches = findUnitMatches(units, this.query, this.options);
			} catch (error) {
				if (!(error instanceof FindPatternError)) throw error;
				this.invalid = true;
			}
		}
		this.activeIndex = this.indexAtSelection();
		return this.status;
	}

	findNext(): SearchMatch | null {
		this.refresh();
		if (!this.getView() || !this.matches.length) return null;
		let next = this.activeIndex + 1;
		if (this.activeIndex < 0) {
			const cursor = this.cursor();
			next = this.matches.findIndex(
				(match) =>
					match.order > cursor.order || (match.order === cursor.order && match.from >= cursor.pos),
			);
			if (next < 0) next = 0;
		}
		return this.select(next % this.matches.length);
	}

	findPrevious(): SearchMatch | null {
		this.refresh();
		if (!this.getView() || !this.matches.length) return null;
		let previous = this.activeIndex - 1;
		if (this.activeIndex < 0) {
			const cursor = this.cursor();
			previous = -1;
			for (let index = this.matches.length - 1; index >= 0; index--) {
				const match = this.matches[index]!;
				if (
					match.order < cursor.order ||
					(match.order === cursor.order && match.from < cursor.pos)
				) {
					previous = index;
					break;
				}
			}
			if (previous < 0) previous = this.matches.length - 1;
		}
		return this.select((previous + this.matches.length) % this.matches.length);
	}

	replaceOne(replacement: string, replaceFormat?: ReplaceFormatPatch): boolean {
		const view = this.getView();
		if (!view?.editable || (!this.query && !hasFormatCriteria(this.options.format))) return false;
		this.refresh();
		if (this.activeIndex < 0 && !this.findNext()) return false;
		const match = this.matches[this.activeIndex];
		if (!match) return false;
		return this.replace(view, [match], replacement, replaceFormat) > 0;
	}

	replaceAll(replacement: string, replaceFormat?: ReplaceFormatPatch): number {
		const view = this.getView();
		if (!view?.editable || (!this.query && !hasFormatCriteria(this.options.format))) return 0;
		this.refresh();
		if (!this.matches.length) return 0;
		if (this.matches.some((match) => match.unit.key !== 'main')) this.stories?.closeStory();
		return this.replace(view, [...this.matches], replacement, replaceFormat);
	}

	/** The open story's selection, a placed match, or the body caret, as a search-order position. */
	private cursor(): { order: number; pos: number } {
		const orderOf = (key: string) => this.units.findIndex((unit) => unit.key === key);
		const open = this.stories?.openStory();
		if (open && orderOf(open.key) >= 0)
			return { order: orderOf(open.key), pos: open.view.state.selection.from };
		const placed = this.currentPlaced();
		if (placed && orderOf(placed.key) >= 0) return { order: orderOf(placed.key), pos: placed.from };
		return { order: 0, pos: this.getView()?.state.selection.from ?? 0 };
	}

	private currentPlaced(): Placed | undefined {
		const selection = this.getView()?.state.selection;
		const placed = this.placed;
		if (!selection || !placed || this.stories?.openStory()) return undefined;
		return selection.from === placed.body.from && selection.to === placed.body.to
			? placed
			: undefined;
	}

	private indexAtSelection(): number {
		const same = (key: string, range: SearchMatch) => (match: StoryMatch) =>
			match.unit.key === key && match.from === range.from && match.to === range.to;
		const open = this.stories?.openStory();
		if (open) return this.matches.findIndex(same(open.key, open.view.state.selection));
		const placed = this.currentPlaced();
		if (placed) return this.matches.findIndex(same(placed.key, placed));
		const selection = this.getView()?.state.selection;
		return selection ? this.matches.findIndex(same('main', selection)) : -1;
	}

	private select(index: number): SearchMatch {
		const view = this.getView()!;
		const match = expectDefined(this.matches[index], 'search match');
		const { unit } = match;
		const place = (body: SearchMatch) => {
			this.placed = { key: unit.key, from: match.from, to: match.to, body };
		};
		this.placed = undefined;
		// Selecting re-enters `refresh` through the editor's change handlers, so the placed match is
		// recorded before the dispatch and the active index is settled again afterwards.
		if (unit.key === 'main' || unit.textBoxPos !== undefined) {
			this.stories?.closeStory();
			const { doc } = view.state;
			let selection: TextSelection | NodeSelection;
			if (unit.textBoxPos === undefined)
				selection = TextSelection.create(doc, match.from, match.to);
			else {
				selection = NodeSelection.create(doc, unit.textBoxPos);
				place({ from: selection.from, to: selection.to });
			}
			view.dispatch(view.state.tr.setSelection(selection).scrollIntoView());
		} else {
			this.stories?.showStory(unit.story, match.from, match.to);
			const { from, to } = view.state.selection;
			if (this.stories?.openStory()?.key !== unit.key) place({ from, to });
		}
		const settled = this.indexAtSelection();
		this.activeIndex = settled < 0 ? index : settled;
		return match;
	}

	/** Replaces the matches, one undoable step per story; the count replaced, 0 if invalid. */
	private replace(
		view: EditorView,
		matches: StoryMatch[],
		replacement: string,
		replaceFormat?: ReplaceFormatPatch,
	): number {
		try {
			const count = applyReplacements(view, this.stories, matches, {
				replacement,
				wildcards: Boolean(this.options.wildcards),
				...(replaceFormat ? { replaceFormat } : {}),
			});
			this.placed = undefined;
			this.refresh();
			return count;
		} catch (error) {
			if (!(error instanceof FindPatternError)) throw error;
			this.invalid = true;
			return 0;
		}
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
