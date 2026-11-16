import { createFindReplaceController } from 'ooxml-ui/pptx';
import type {
	FindReplaceController,
	FindReplaceMode,
	FindReplaceOptions,
	FindResult,
} from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

import { controllerState } from '../state/controller-state';
import type { ControllerState } from '../state/controller-state';

/**
 * Reactive Find & Replace panel state for the Home tab's Editing group: the
 * shared `createFindReplaceController` (the bar's open flag and mode, and the
 * search session) read through `controllerState`, with every replace routed
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
	readonly controller: FindReplaceController;
	readonly #state: ControllerState<ReturnType<FindReplaceController['getSnapshot']>>;

	constructor(deps: FindReplaceDeps) {
		this.controller = createFindReplaceController({
			slides: () => deps.getSlides(),
			commit: (next) => deps.commitSlides([...next]),
			navigate: ({ slideIndex, elementId }) => deps.onNavigate?.(slideIndex, elementId),
		});
		this.#state = controllerState(this.controller);
	}

	get #session() {
		return this.#state.current.session;
	}

	/** Whether the bar is open; setting it opens (searching again) or closes it. */
	get open(): boolean {
		return this.#state.current.open;
	}
	set open(open: boolean) {
		this.controller.setOpen(open);
	}

	/** Find (Ctrl+F) or Find and Replace (Ctrl+H). */
	get mode(): FindReplaceMode {
		return this.#state.current.mode;
	}

	get query(): string {
		return this.#session.query;
	}
	set query(query: string) {
		this.controller.search({ query });
	}

	get replacement(): string {
		return this.#session.replacement;
	}
	set replacement(replacement: string) {
		this.controller.setReplacement(replacement);
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

	/** Ctrl+F / Ctrl+H: close the bar open in `mode`, switch its mode, or open it. */
	toggle(mode?: FindReplaceMode): void {
		this.controller.toggle(mode);
	}

	close(): void {
		this.controller.close();
	}

	/** Re-run the search from the current query and scopes and jump to the first match. */
	search(): void {
		this.controller.search();
	}

	next(): void {
		this.controller.next();
	}

	prev(): void {
		this.controller.previous();
	}

	/** Replace the current match, then move on to the next one. */
	replaceCurrent(): void {
		this.controller.replaceCurrent();
	}

	/** Replace every match of the current query, then search again. */
	replaceAll(): void {
		this.controller.replaceAll();
	}

	#setOption(options: Partial<FindReplaceOptions>): void {
		this.controller.search({ options });
	}
}
