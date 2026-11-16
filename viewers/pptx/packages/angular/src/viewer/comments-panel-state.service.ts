import { DestroyRef, Injectable, effect, inject, signal } from '@angular/core';
import type { Signal } from '@angular/core';
import { createCommentsPanelController, INITIAL_COMMENTS_PANEL_STATE } from 'ooxml-ui/pptx';
import type { CommentsPanelController, CommentsPanelState } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

import { EditorStateService } from './editor-state.service';

/** What the viewer supplies once its own signals exist. */
export interface CommentsPanelHost {
	activeSlideIndex: () => number;
	canEdit: () => boolean;
	authorName: () => string;
	selectedElementId: () => string | undefined;
}

/**
 * The viewer's comments panel state: one shared `createCommentsPanelController` (drafts, edits and
 * reply composers, per slide) that every `<pptx-comments-panel>` under the viewer reads, whether
 * it sits in the side rail or the inspector's Comments tab. Provided per viewer
 * (`POWER_POINT_VIEWER_PROVIDERS`); the viewer calls {@link attach} with its signals.
 */
@Injectable()
export class CommentsPanelStateService {
	private readonly editor = inject(EditorStateService);
	private readonly host = signal<CommentsPanelHost | null>(null);
	private readonly snapshot = signal<CommentsPanelState>(INITIAL_COMMENTS_PANEL_STATE);

	readonly controller: CommentsPanelController = createCommentsPanelController({
		slides: () => this.editor.slides(),
		canEdit: () => this.host()?.canEdit() ?? false,
		authorName: () => this.host()?.authorName() ?? 'You',
		selectedElementId: () => this.host()?.selectedElementId(),
		// The controller changes one slide's comments per command: write that slide, one
		// history entry, through the editor's own slide patch.
		commit: (next: PptxSlide[]) => {
			const current = this.editor.slides();
			const index = next.findIndex((slide, i) => slide !== current[i]);
			const slide = next[index];
			if (slide) {
				this.editor.updateSlide(index, { comments: slide.comments ?? [] });
			}
		},
	});
	/** The controller's snapshot as a signal. */
	readonly state: Signal<CommentsPanelState> = this.snapshot.asReadonly();

	constructor() {
		const unsubscribe = this.controller.subscribe(() =>
			this.snapshot.set(this.controller.getSnapshot()),
		);
		inject(DestroyRef).onDestroy(unsubscribe);
		// Drafts for slides that no longer exist go with them.
		effect(() => {
			this.editor.slides();
			this.controller.pruneDrafts();
		});
	}

	attach(host: CommentsPanelHost): void {
		this.host.set(host);
	}

	slideIndex(): number {
		return this.host()?.activeSlideIndex() ?? 0;
	}

	slideId(): string | undefined {
		return this.editor.slides()[this.slideIndex()]?.id;
	}

	canEdit(): boolean {
		return this.host()?.canEdit() ?? false;
	}
}
