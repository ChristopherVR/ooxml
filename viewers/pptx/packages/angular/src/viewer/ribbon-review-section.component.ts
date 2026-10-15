import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	computed,
	DestroyRef,
	inject,
	Input,
	output,
	signal,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import {
	buildReviewRibbon,
	commentCounts,
	findCommentLocation,
	isReviewCommentCommand,
	runReviewCommentCommand,
	trackActiveComment,
} from 'ooxml-ui/pptx';
import type { ReviewCommentsHost, RibbonCommandRequestEvent } from 'ooxml-ui/pptx';

import { EditorStateService } from './editor-state.service';
import { ReviewCommentsNavigationService } from './review-comments-navigation.service';
import { ViewerInspectorPanelService } from './viewer-inspector-panel.service';

@Component({
	selector: 'pptx-ribbon-review-section',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `<pptx-ui-ribbon-section [groups]="groups()" (command-request)="request($event)" />`,
})
export class RibbonReviewSectionComponent {
	private readonly translator = inject(TranslateService);
	private readonly editor = inject(EditorStateService, { optional: true });
	private readonly nav = inject(ReviewCommentsNavigationService, { optional: true });
	private readonly panel = inject(ViewerInspectorPanelService, { optional: true });
	/** The comment Previous / Next / Delete act on; follows the user's own picks in the panel. */
	private readonly activeCommentId = signal<string | null>(null);
	private readonly commentsHost: ReviewCommentsHost | null =
		this.editor && this.nav
			? {
					slides: () => this.editor?.slides() ?? [],
					currentSlide: () => this.nav?.currentSlide() ?? 0,
					goToSlide: (index) => this.nav?.goToSlide(index),
					ensureCommentsOpen: () => {
						this.panel?.mobileInspectorHidden.set(false);
						this.panel?.activePanel.set('comments');
					},
					applyComments: (changes) => {
						const editor = this.editor;
						if (!editor) {
							return;
						}
						const next = editor.slides().map((slide, index) => {
							const change = changes.find((candidate) => candidate.slideIndex === index);
							return change ? { ...slide, comments: change.comments } : slide;
						});
						editor.applyReplacement(next);
					},
				}
			: null;
	constructor() {
		const stop = trackActiveComment(document, (id) => this.activeCommentId.set(id));
		inject(DestroyRef).onDestroy(stop);
	}
	private readonly state = signal({ spellCheck: false, editable: false });
	@Input() get spellCheckEnabled(): boolean {
		return this.state().spellCheck;
	}
	set spellCheckEnabled(value: boolean) {
		this.state.update((state) => ({ ...state, spellCheck: value }));
	}
	@Input() get canEdit(): boolean {
		return this.state().editable;
	}
	set canEdit(value: boolean) {
		this.state.update((state) => ({ ...state, editable: value }));
	}
	readonly comments = output<void>();
	readonly spellCheckChange = output<boolean>();
	readonly a11y = output<void>();
	readonly openCompare = output<void>();
	readonly language = output<void>();
	readonly link = output<void>();
	readonly groups = computed(() =>
		buildReviewRibbon((key) => this.translator.instant(key) as string, {
			editable: this.canEdit,
			spellCheck: this.spellCheckEnabled,
			canAccessibility: true,
			canLanguage: true,
			canCompare: true,
			canComments: true,
			...this.commentState(),
		}),
	);
	private commentState(): {
		commentCount: number;
		totalCommentCount: number;
		canNavigateComments: boolean;
		hasActiveComment: boolean;
	} {
		const slides = this.editor?.slides() ?? [];
		const counts = commentCounts(slides, this.nav?.currentSlide() ?? 0);
		return {
			commentCount: counts.slide,
			totalCommentCount: counts.total,
			canNavigateComments: this.commentsHost !== null,
			hasActiveComment: findCommentLocation(slides, this.activeCommentId()) !== null,
		};
	}
	protected request(event: Event): void {
		const id = (event as RibbonCommandRequestEvent).detail.id;
		if (isReviewCommentCommand(id)) {
			const navigation = id === 'review.comments.previous' || id === 'review.comments.next';
			if (this.commentsHost && (navigation || this.canEdit)) {
				const next = runReviewCommentCommand(id, this.commentsHost, this.activeCommentId());
				if (next !== undefined) {
					this.activeCommentId.set(next);
				}
			}
			return;
		}
		switch (id) {
			case 'review.proofing.spelling':
				this.spellCheckChange.emit(!this.spellCheckEnabled);
				break;
			case 'review.accessibility.check':
				this.a11y.emit();
				break;
			case 'review.language.language':
				this.language.emit();
				break;
			case 'review.compare.compare':
				if (this.canEdit) {
					this.openCompare.emit();
				}
				break;
			case 'review.comments.newComment':
			case 'review.comments.showComments':
				this.comments.emit();
				break;
		}
	}
}
