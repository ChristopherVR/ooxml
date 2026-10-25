/**
 * viewer-find-replace.service.ts: Viewer-scoped state for the find-in-slides bar
 * and the edit-mode find-and-replace bar. Owns the two bar visibility flags and
 * keeps the shared Find and Replace session (`find-replace-session`: matches,
 * cursor, replace plans, the search after a replace) in a signal.
 *
 * Extracted from {@link PowerPointViewerComponent}: the component binds a slide
 * navigation callback via {@link bind} (so a match can bring its slide into
 * view, and the editor selects its element) and the template reads the flags /
 * invokes the handlers off the injected instance.
 *
 * Provide it once on the viewer component (`providers: [ViewerFindReplaceService]`).
 */

import { computed, inject, Injectable, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import {
	DEFAULT_FIND_REPLACE_OPTIONS,
	findReplaceSessionActions,
	INITIAL_FIND_REPLACE_SESSION,
} from 'ooxml-ui/pptx';
import type { FindReplaceSession } from 'ooxml-ui/pptx';

import { EditorStateService } from './editor-state.service';
import type { FindEvent, ReplaceEvent } from './find-replace-bar.component';

@Injectable()
export class ViewerFindReplaceService {
	private readonly editor = inject(EditorStateService);
	private readonly translate = inject(TranslateService);

	/** Find-in-slides bar visibility. */
	readonly showFind = signal(false);
	/** Find-and-replace bar visibility (edit mode only). */
	readonly showFindReplace = signal(false);
	/** The find-and-replace bar's session; the bar's fields are bound to it. */
	readonly session = signal<FindReplaceSession>(INITIAL_FIND_REPLACE_SESSION);
	readonly results = computed(() => this.session().matches);
	/** Zero-based current match; -1 when there is none. */
	readonly activeIndex = computed(() => this.session().current);

	/** Navigate the viewer to a slide index (bound from the host component). */
	private goTo: (index: number) => void = () => undefined;

	private readonly actions = findReplaceSessionActions(
		{ get: () => this.session(), set: (next) => this.session.set(next) },
		{
			slides: () => this.editor.slides(),
			commit: (slides, scope) =>
				this.editor.applyReplacement(
					slides,
					this.translate.instant(
						scope === 'all' ? 'pptx.undoAction.replaceAll' : 'pptx.undoAction.replace',
					),
				),
			navigate: ({ slideIndex, elementId }) => {
				this.goTo(slideIndex);
				this.editor.select(elementId ? [elementId] : []);
			},
		},
	);

	/** Wire the host's slide-navigation callback (called once from the constructor). */
	bind(goTo: (index: number) => void): void {
		this.goTo = goTo;
	}

	/** Open the find/replace bar (mutually exclusive with the find-only bar). */
	openFindReplace(): void {
		const wasOpen = this.showFindReplace();
		this.showFind.set(false);
		this.showFindReplace.set(true);
		if (!wasOpen) {
			// The bar shows the query the session kept; search it again from the first match.
			this.actions.search();
		}
	}

	/** Close the find/replace bar; its matches go with it. */
	closeFindReplace(): void {
		this.showFindReplace.set(false);
		this.actions.close();
	}

	onFind(evt: FindEvent): void {
		this.actions.search({
			query: evt.query,
			options: {
				matchCase: evt.matchCase,
				includeNotes: evt.includeNotes ?? DEFAULT_FIND_REPLACE_OPTIONS.includeNotes,
				includeChartTitles:
					evt.includeChartTitles ?? DEFAULT_FIND_REPLACE_OPTIONS.includeChartTitles,
				includeChartData: evt.includeChartData ?? DEFAULT_FIND_REPLACE_OPTIONS.includeChartData,
			},
		});
	}

	onReplacementChange(replacement: string): void {
		this.actions.setReplacement(replacement);
	}

	onNavigate(dir: 1 | -1): void {
		if (dir > 0) {
			this.actions.next();
		} else {
			this.actions.previous();
		}
	}

	onReplaceOne(evt: ReplaceEvent): void {
		this.actions.setReplacement(evt.replacement);
		this.actions.replaceCurrent();
	}

	onReplaceAll(evt: ReplaceEvent): void {
		this.actions.setReplacement(evt.replacement);
		this.actions.replaceAll();
	}
}
