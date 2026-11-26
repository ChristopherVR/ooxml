/**
 * comments-panel.component.ts: Side panel listing the active slide's comments.
 *
 * Selector: `pptx-comments-panel`
 *
 * It renders the supplied `comments` (the active slide's, nested replies included). The drafts,
 * the comment being edited, the open reply composer and every command are the viewer's shared
 * comments panel controller (`CommentsPanelStateService`), so a half-typed comment stays with its
 * slide and the side rail and the inspector's Comments tab show the same drafts.
 *
 * Usage:
 * ```html
 * <pptx-comments-panel [comments]="slideComments()" [authorName]="userName()" />
 * ```
 */

import {
	afterNextRender,
	ChangeDetectionStrategy,
	Component,
	computed,
	DestroyRef,
	ElementRef,
	inject,
	input,
	output,
} from '@angular/core';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import {
	attachCommentsTaskPane,
	createCommentsPanelController,
	INITIAL_COMMENTS_PANEL_STATE,
} from 'ooxml-ui/pptx';
import type { CommentsPanelController, CommentsPanelState } from 'ooxml-ui/pptx';
import type { PptxComment, PptxCommentMention, PptxModernCommentAuthor } from 'pptx-viewer-core';
import { formatCommentTimestamp } from 'pptx-viewer-core';

import { CommentBodyComponent } from './comment-body.component';
import { CommentMentionTextareaComponent } from './comment-mention-textarea.component';
import { CommentsPanelStateService } from './comments-panel-state.service';

export interface CommentSubmission {
	text: string;
	mentions: PptxCommentMention[];
}

@Component({
	selector: 'pptx-comments-panel',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [TranslatePipe, CommentBodyComponent, CommentMentionTextareaComponent],
	templateUrl: './comments-panel.component.html',
	styleUrl: './comments-panel.component.css',
})
export class CommentsPanelComponent {
	/** The active slide's comments (already filtered to the active slide). */
	readonly comments = input<PptxComment[]>([]);

	private readonly translate = inject(TranslateService);
	readonly closeable = input(false);
	readonly closed = output<void>();

	constructor() {
		const host = inject<ElementRef<HTMLElement>>(ElementRef);
		const destroy = inject(DestroyRef);
		afterNextRender(() => {
			if (this.closeable()) {
				destroy.onDestroy(
					attachCommentsTaskPane(host.nativeElement, {
						label: this.translate.instant('pptx.toolbar.comments'),
						closeLabel: this.translate.instant('pptx.common.close'),
						close: () => this.closed.emit(),
					}),
				);
			}
		});
	}

	/** Display name shown in the compose label ("Commenting as ..."). */
	readonly authorName = input<string>(this.translate.instant('pptx.comments.defaultAuthorName'));

	/** Office 2021 modern comment authors, for the `@`-mention typeahead. */
	readonly modernCommentAuthors = input<readonly PptxModernCommentAuthor[]>([]);

	/**
	 * The viewer's shared comments panel state (drafts, edits and reply composers, per slide).
	 * Outside a viewer there is no deck to write to, so the panel lists the comments read-only.
	 */
	private readonly panel = inject(CommentsPanelStateService, { optional: true });
	protected readonly controller: CommentsPanelController =
		this.panel?.controller ??
		createCommentsPanelController({
			slides: () => [],
			canEdit: () => false,
			authorName: () => '',
			commit: () => undefined,
		});
	private readonly state = computed<CommentsPanelState>(
		() => this.panel?.state() ?? INITIAL_COMMENTS_PANEL_STATE,
	);
	protected readonly slideId = computed(() => this.panel?.slideId() ?? '');
	protected readonly canEdit = computed(() => this.panel?.canEdit() ?? false);
	private slideIndex(): number {
		return this.panel?.slideIndex() ?? 0;
	}

	/** The active slide's new-comment draft, as typed. */
	readonly draft = computed(() => this.state().draftBySlideId[this.slideId()] ?? '');
	/** `@`-mentions recorded in {@link draft} so far. */
	readonly draftMentions = computed(
		() => (this.state().draftMentionsBySlideId[this.slideId()] ?? []) as PptxCommentMention[],
	);
	/** Whether the draft has non-whitespace content (enables the submit button). */
	readonly canAdd = computed<boolean>(() => this.canEdit() && this.draft().trim().length > 0);

	/** The comment being edited on this slide, or null. */
	readonly editingId = computed(() => this.state().editingIdBySlideId[this.slideId()] ?? null);
	/** Id of the comment whose reply composer is open (one at a time). */
	readonly replyingTo = computed(() => this.state().replyingToId);

	protected setDraft(text: string): void {
		this.controller.setDraft(this.slideId(), text);
	}

	protected setDraftMentions(mentions: PptxCommentMention[]): void {
		this.controller.setDraft(this.slideId(), this.draft(), mentions);
	}

	submit(event: Event): void {
		event.preventDefault();
		this.controller.add(this.slideIndex());
	}

	protected removeComment(id: string): void {
		this.controller.remove(this.slideIndex(), id);
	}

	protected toggleResolved(id: string): void {
		this.controller.toggleResolved(this.slideIndex(), id);
	}

	protected startEdit(id: string): void {
		this.controller.startEdit(this.slideId(), id);
	}

	protected cancelEdit(): void {
		this.controller.cancelEdit(this.slideId());
	}

	protected editDraft(id: string): string {
		return this.state().editDraftById[id] ?? '';
	}

	protected setEditDraft(id: string, event: Event): void {
		this.controller.setEditDraft(id, (event.target as HTMLTextAreaElement).value);
	}

	protected saveEdit(id: string): void {
		this.controller.saveEdit(this.slideIndex(), id);
	}

	/** Localized reply-composer placeholder ("Reply to <author>..."). */
	replyPlaceholder(comment: PptxComment): string {
		return this.translate.instant('pptx.comments.replyPlaceholder', {
			author: comment.author || this.translate.instant('pptx.comments.unknownAuthor'),
		});
	}

	/** Open the reply composer under `parentId` (closing any other composer), empty. */
	startReply(parentId: string): void {
		this.controller.startReply(parentId);
	}

	cancelReply(): void {
		this.controller.cancelReply();
	}

	protected replyDraft(id: string): string {
		return this.state().replyDraftById[id] ?? '';
	}

	protected replyMentions(id: string): PptxCommentMention[] {
		return (this.state().replyMentionsById[id] ?? []) as PptxCommentMention[];
	}

	protected setReplyDraft(id: string, text: string): void {
		this.controller.setReplyDraft(id, text);
	}

	protected setReplyMentions(id: string, mentions: PptxCommentMention[]): void {
		this.controller.setReplyDraft(id, this.replyDraft(id), mentions);
	}

	submitReply(parentId: string): void {
		this.controller.submitReply(this.slideIndex(), parentId);
	}

	formatTimestamp(value: string | undefined): string {
		return formatCommentTimestamp(value);
	}
}
