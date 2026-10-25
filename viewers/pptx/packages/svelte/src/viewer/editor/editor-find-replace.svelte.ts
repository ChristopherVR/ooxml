import { findReplaceSessionActions, INITIAL_FIND_REPLACE_SESSION } from 'ooxml-ui/pptx';
import type {
	FindReplaceOptions,
	FindReplaceSession,
	FindReplaceSessionActions,
	FindResult,
} from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

/**
 * Reactive Find & Replace panel state for the Home tab's Editing group.
 * The session (matches, current match, replace plans, the search after a
 * replace) is the shared `find-replace-session` model, kept here in
 * `$state.raw`; this class adds the panel's open flag and routes every replace
 * through the host's history-tracked commit so undo/redo cover it.
 */
export interface FindReplaceDeps {
	getSlides(): readonly PptxSlide[];
	/** Commit a fully-replaced slide array (pushes history, marks dirty). */
	commitSlides(next: PptxSlide[]): void;
	/** Called with the match's slide and element (null for a speaker-notes match). */
	onNavigate?(slideIndex: number, elementId: string | null): void;
}

export class FindReplaceState {
	open = $state(false);
	#session = $state.raw<FindReplaceSession>(INITIAL_FIND_REPLACE_SESSION);
	readonly #actions: FindReplaceSessionActions;

	constructor(deps: FindReplaceDeps) {
		this.#actions = findReplaceSessionActions(
			{
				get: () => this.#session,
				set: (next) => {
					this.#session = next;
				},
			},
			{
				slides: () => deps.getSlides(),
				commit: (next) => deps.commitSlides([...next]),
				navigate: ({ slideIndex, elementId }) => deps.onNavigate?.(slideIndex, elementId),
			},
		);
	}

	/** The query; setting it searches. */
	get query(): string {
		return this.#session.query;
	}
	set query(query: string) {
		this.#actions.search({ query });
	}

	get replacement(): string {
		return this.#session.replacement;
	}
	set replacement(replacement: string) {
		this.#actions.setReplacement(replacement);
	}

	get matchCase(): boolean {
		return this.#session.options.matchCase;
	}
	set matchCase(matchCase: boolean) {
		this.#setOption({ matchCase });
	}

	get includeNotes(): boolean {
		return this.#session.options.includeNotes;
	}
	set includeNotes(includeNotes: boolean) {
		this.#setOption({ includeNotes });
	}

	get includeChartTitles(): boolean {
		return this.#session.options.includeChartTitles;
	}
	set includeChartTitles(includeChartTitles: boolean) {
		this.#setOption({ includeChartTitles });
	}

	get includeChartData(): boolean {
		return this.#session.options.includeChartData;
	}
	set includeChartData(includeChartData: boolean) {
		this.#setOption({ includeChartData });
	}

	get results(): readonly FindResult[] {
		return this.#session.matches;
	}

	/** Zero-based current match; -1 when there is none. */
	get index(): number {
		return this.#session.current;
	}

	get hasResults(): boolean {
		return this.results.length > 0;
	}

	get matchCount(): number {
		return this.results.length;
	}

	toggle(): void {
		if (this.open) {
			this.close();
			return;
		}
		this.open = true;
		this.search();
	}

	/** Close the panel; its matches go with it. */
	close(): void {
		this.open = false;
		this.#actions.close();
	}

	/** Re-run the search from the current query and scopes and jump to the first match. */
	search(): void {
		this.#actions.search();
	}

	next(): void {
		this.#actions.next();
	}

	prev(): void {
		this.#actions.previous();
	}

	/** Replace the current match, then move on to the next one. */
	replaceCurrent(): void {
		this.#actions.replaceCurrent();
	}

	/** Replace every match of the current query, then search again. */
	replaceAll(): void {
		this.#actions.replaceAll();
	}

	#setOption(options: Partial<FindReplaceOptions>): void {
		this.#actions.search({ options });
	}
}
