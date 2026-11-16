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

import { computed, inject, Injectable } from '@angular/core';
import type { Signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { createFindReplaceController, DEFAULT_FIND_REPLACE_OPTIONS } from 'ooxml-ui/pptx';
import type { FindReplaceMode, FindReplaceSession } from 'ooxml-ui/pptx';

import { controllerSignal } from './controller-signal';
import { EditorStateService } from './editor-state.service';
import type { FindEvent, ReplaceEvent } from './find-replace-bar.component';

@Injectable()
export class ViewerFindReplaceService {
	private readonly editor = inject(EditorStateService);
	private readonly translate = inject(TranslateService);

	/** Find-in-slides bar visibility. */
	private goTo: (index: number) => void = () => undefined;

	/** The bar's open state and mode plus the search session: the shared controller. */
	private readonly actions = createFindReplaceController({
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
	});
	private readonly state = controllerSignal(this.actions);

	/** Whether the bar is open (either mode). */
	readonly isOpen: Signal<boolean> = computed(() => this.state().open);
	/** Find mode (Ctrl+F). */
	readonly showFind: Signal<boolean> = computed(
		() => this.state().open && this.state().mode === 'find',
	);
	/** Find and Replace mode (Ctrl+H). */
	readonly showFindReplace: Signal<boolean> = computed(
		() => this.state().open && this.state().mode === 'replace',
	);
	readonly session: Signal<FindReplaceSession> = computed(() => this.state().session);
	readonly results = computed(() => this.session().matches);
	readonly activeIndex = computed(() => this.session().current);

	bind(goTo: (index: number) => void): void {
		this.goTo = goTo;
	}

	/** Open the bar in `mode` (searching the kept query again), or switch an open bar's mode. */
	open(mode: FindReplaceMode = 'find'): void {
		this.actions.open(mode);
	}

	openFindReplace(): void {
		this.actions.open('replace');
	}

	/** Close the bar: the matches go, the query, replacement and scopes stay. */
	close(): void {
		this.actions.close();
	}

	closeFindReplace(): void {
		this.actions.close();
	}

	/** Ctrl+F / Ctrl+H and the toolbar: close the bar open in `mode`, switch, or open it. */
	toggle(mode: FindReplaceMode = 'find'): void {
		this.actions.toggle(mode);
	}

	/** The deck changed (a replace searches again once the replaced slides land). */
	slidesChanged(): void {
		this.actions.slidesChanged();
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
