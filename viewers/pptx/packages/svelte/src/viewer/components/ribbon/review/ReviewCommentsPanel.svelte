<script lang="ts">
	/**
	 * ReviewCommentsPanel: history-aware comment review for the active slide.
	 * The drafts, the comment being edited, the open reply composer and every
	 * command are the viewer's shared comments panel controller
	 * (`commentsPanelFor`), so a half-typed comment stays with its slide; this
	 * component only owns the compact review UI.
	 *
	 * Wave-4 B5: both composers (new comment, reply) offer an `@`-mention
	 * typeahead ({@link CommentComposeState} + {@link CommentMentionSuggestions}),
	 * and replies render recursively ({@link CommentReplyThread}) so a legacy
	 * `p:cmLst` reply nested more than one level deep is no longer dropped.
	 *
	 * The accepted mentions index the draft as typed; `addCommentToList` /
	 * `replyToCommentInList` (shared) store them moved with the draft's trim.
	 */
	import type { PptxModernCommentAuthor } from 'pptx-viewer-core';
	import { untrack } from 'svelte';

	import { useTranslator } from '../../../../i18n/context';
	import { useViewerOptions } from '../../../state/viewer-options-context';
	import CommentBody from '../../CommentBody.svelte';
	import type { EditorState } from '../../../editor/editor-state.svelte';
	import { controllerState } from '../../../state/controller-state';
	import { CommentComposeState } from './comment-compose.svelte';
	import { commentsPanelFor } from './comments-panel';
	import CommentMentionSuggestions from './CommentMentionSuggestions.svelte';
	import CommentReplyThread from './CommentReplyThread.svelte';

	const {
		editor,
		embedded = false,
	}: {
		editor: EditorState;
		/**
		 * True when hosted inside a chrome that already renders its own title +
		 * close button (the mobile `MobileSheet`). Suppresses the internal
		 * heading so mobile doesn't show "Comments" twice stacked.
		 */
		embedded?: boolean;
	} = $props();
	const t = useTranslator();
	const optionsState = useViewerOptions();
	/** Host `authorName`, then Options > General > "User name", then the generic default label. */
	const authorName = $derived(
		editor.hostAuthorName ||
			optionsState.options.general.userName ||
			t('pptx.comments.defaultAuthorName'),
	);
	// One viewer, one controller: the editor prop never changes for a mounted panel.
	const controller = commentsPanelFor(untrack(() => editor), () => authorName);
	const panel = controllerState(controller);
	const slide = $derived(editor.slides[editor.currentSlideIndex]);
	const slideId = $derived(slide?.id ?? '');
	// Both composers keep their drafts in the controller: the new comment per slide, the reply
	// per parent comment.
	const compose = new CommentComposeState({
		read: () => ({
			text: panel.current.draftBySlideId[slideId] ?? '',
			mentions: panel.current.draftMentionsBySlideId[slideId] ?? [],
		}),
		write: (text, mentions) => controller.setDraft(slideId, text, mentions),
	});
	const replyingTo = $derived(panel.current.replyingToId);
	const reply = new CommentComposeState({
		read: () => ({
			text: replyingTo ? (panel.current.replyDraftById[replyingTo] ?? '') : '',
			mentions: replyingTo ? (panel.current.replyMentionsById[replyingTo] ?? []) : [],
		}),
		write: (text, mentions) => {
			if (replyingTo) {
				controller.setReplyDraft(replyingTo, text, mentions);
			}
		},
	});
	const editingId = $derived(panel.current.editingIdBySlideId[slideId] ?? null);
	$effect(() => {
		void editor.slides;
		controller.pruneDrafts();
	});
	const comments = $derived(slide?.comments ?? []);
	const selectedLabel = $derived(editor.selectedElement?.type ?? null);
	/**
	 * The `@`-mention author catalogue: modern authors plus legacy
	 * (`ppt/commentAuthors.xml`) ones mapped into the same shape (their `id`
	 * doubles as `userId`; `providerId` is a fixed marker, unused by matching
	 * or by `insertCommentMention`, which only reads `id`/`name`).
	 */
	const mentionAuthors = $derived<PptxModernCommentAuthor[]>([
		...editor.modernCommentAuthors,
		...editor.commentAuthors.map((author) => ({
			id: author.id,
			name: author.name,
			initials: author.initials,
			userId: author.id,
			providerId: 'legacy',
		})),
	]);

	function addComment(): void {
		controller.add(editor.currentSlideIndex);
	}

	function toggleResolved(id: string): void {
		controller.toggleResolved(editor.currentSlideIndex, id);
	}

	function removeComment(id: string): void {
		controller.remove(editor.currentSlideIndex, id);
	}

	function startReply(id: string): void {
		controller.startReply(id);
		reply.reset();
	}

	function cancelReply(): void {
		controller.cancelReply();
	}

	function submitReply(id: string): void {
		controller.submitReply(editor.currentSlideIndex, id);
	}

	// Wire a composer's `<textarea>` to its `CommentComposeState` for the
	// typeahead: caret tracking on input/click/arrow-keys, and the suggestion
	// list's keyboard navigation. Two small factories (not one that returns a
	// handler bag) so each stays a plain, directly-typed element attribute --
	// the idiom this codebase uses everywhere else for a typed `currentTarget`.
	function onComposeChange(state: CommentComposeState) {
		return (event: Event & { currentTarget: HTMLTextAreaElement }) => {
			state.onInput(event.currentTarget.value, event.currentTarget.selectionStart ?? 0);
		};
	}
	function onComposeKeydown(state: CommentComposeState) {
		return (event: KeyboardEvent & { currentTarget: HTMLTextAreaElement }) => {
			const result = state.onKeydown(event, mentionAuthors);
			if (!result.consumed) {
				return;
			}
			event.preventDefault();
			if (result.caret === undefined) {
				return;
			}
			const target = event.currentTarget;
			const caret = result.caret;
			queueMicrotask(() => target.setSelectionRange(caret, caret));
		};
	}
</script>

<section
	class="pptx-svelte-comments"
	aria-label={embedded ? t('pptx.comments.slideComments') : undefined}
	aria-labelledby={embedded ? undefined : 'pptx-svelte-comments-title'}
>
	{#if !embedded}
		<div class="pptx-svelte-comments-heading">
			<div>
				<span class="pptx-svelte-comments-eyebrow">{t('pptx.ribbon.tab.review')}</span>
				<h3 id="pptx-svelte-comments-title">{t('pptx.comments.slideComments')}</h3>
			</div>
			<span class="pptx-svelte-comments-count">{comments.length}</span>
		</div>
	{/if}

	{#if selectedLabel}
		<p class="pptx-svelte-comments-target">{t('pptx.comments.commentingOn')} {selectedLabel}</p>
	{/if}
	<div class="pptx-svelte-comments-compose">
		<div class="pptx-svelte-comment-composer">
			<textarea
				data-pptx-comment-draft
				value={compose.text}
				rows="2"
				placeholder={t('pptx.comments.mentionPlaceholder')}
				aria-label={t('pptx.comments.addComment')}
				oninput={onComposeChange(compose)}
				onkeyup={onComposeChange(compose)}
				onclick={onComposeChange(compose)}
				onkeydown={onComposeKeydown(compose)}
			></textarea>
			<CommentMentionSuggestions
				authors={compose.suggestions(mentionAuthors)}
				highlightIndex={compose.highlightIndex}
				onselect={(author) => compose.accept(author)}
			/>
		</div>
		<button type="button" disabled={!editor.editable || !compose.text.trim()} onclick={addComment}>{t('pptx.comments.addComment')}</button>
	</div>

	{#if comments.length === 0}
		<p class="pptx-svelte-comments-empty">{t('pptx.comments.noneOnSlide')}</p>
	{:else}
		<div class="pptx-svelte-comments-list" aria-label={t('pptx.comments.slideComments')}>
			{#each comments as comment (comment.id)}
				<article class:resolved={comment.resolved} class="pptx-svelte-comment-card" data-comment-id={comment.id}>
					<div class="pptx-svelte-comment-meta">
						<strong>{comment.author ?? t('pptx.comments.unknownAuthor')}</strong>
						{#if comment.resolved}<span>{t('pptx.comments.resolved')}</span>{/if}
					</div>
					{#if editingId === comment.id}
						<div class="pptx-svelte-comment-reply-compose">
							<textarea
								class="pptx-svelte-comment-edit"
								rows="2"
								aria-label={t('pptx.comments.edit')}
								value={panel.current.editDraftById[comment.id] ?? ''}
								oninput={(event) => controller.setEditDraft(comment.id, event.currentTarget.value)}
							></textarea>
							<div class="pptx-svelte-comment-reply-buttons">
								<button type="button" class="pptx-svelte-comment-reply-cancel" onclick={() => controller.cancelEdit(slideId)}>{t('pptx.comments.cancel')}</button>
								<button type="button" class="pptx-svelte-comment-reply-submit" disabled={!(panel.current.editDraftById[comment.id] ?? '').trim()} onclick={() => controller.saveEdit(editor.currentSlideIndex, comment.id)}>{t('pptx.comments.save')}</button>
							</div>
						</div>
					{:else}
						<p><CommentBody text={comment.text} mentions={comment.mentions} /></p>
					{/if}
					<CommentReplyThread replies={comment.replies ?? []} />
					<div class="pptx-svelte-comment-actions">
						<button type="button" onclick={() => toggleResolved(comment.id)}>{comment.resolved ? t('pptx.comments.reopen') : t('pptx.comments.resolve')}</button>
						<button type="button" onclick={() => startReply(comment.id)}>{t('pptx.comments.reply')}</button>
						{#if editor.editable && editingId !== comment.id}<button type="button" class="pptx-svelte-comment-edit-button" onclick={() => controller.startEdit(slideId, comment.id)}>{t('pptx.comments.edit')}</button>{/if}
						<button type="button" onclick={() => removeComment(comment.id)}>{t('pptx.comments.remove')}</button>
					</div>
					{#if replyingTo === comment.id}
						<div class="pptx-svelte-comment-reply-compose">
							<div class="pptx-svelte-comment-composer">
								<textarea
									value={reply.text}
									rows="2"
									placeholder={t('pptx.comments.replyPlaceholder', { author: comment.author ?? t('pptx.comments.unknownAuthor') })}
									aria-label={t('pptx.comments.reply')}
									oninput={onComposeChange(reply)}
								onkeyup={onComposeChange(reply)}
								onclick={onComposeChange(reply)}
								onkeydown={onComposeKeydown(reply)}
								></textarea>
								<CommentMentionSuggestions
									authors={reply.suggestions(mentionAuthors)}
									highlightIndex={reply.highlightIndex}
									onselect={(author) => reply.accept(author)}
								/>
							</div>
							<div class="pptx-svelte-comment-reply-buttons">
								<button type="button" class="pptx-svelte-comment-reply-cancel" onclick={cancelReply}>{t('pptx.comments.cancel')}</button>
								<button type="button" class="pptx-svelte-comment-reply-submit" disabled={!reply.text.trim()} onclick={() => submitReply(comment.id)}>{t('pptx.comments.reply')}</button>
							</div>
						</div>
					{/if}
				</article>
			{/each}
		</div>
	{/if}
</section>

<style>
	.pptx-svelte-comments { display: grid; gap: 8px; width: min(340px, 100%); padding-left: 12px; border-left: 1px solid var(--pptx-border, var(--office-border, #d1d5db)); }
	.pptx-svelte-comments-heading { display: flex; justify-content: space-between; align-items: flex-start; gap: 10px; }
	.pptx-svelte-comments-eyebrow { display: block; color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280)); font-size: 10px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
	.pptx-svelte-comments h3 { margin: 1px 0 0; font-size: 13px; }
	.pptx-svelte-comments-count { display: grid; place-items: center; min-width: 20px; height: 20px; border-radius: 10px; background: var(--pptx-muted, var(--office-surface, #f3f4f6)); font-size: 11px; font-weight: 700; }
	.pptx-svelte-comments-target, .pptx-svelte-comments-empty { margin: 0; color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280)); font-size: 11px; }
	.pptx-svelte-comments-compose { display: grid; gap: 5px; }
	.pptx-svelte-comment-composer { position: relative; }
	.pptx-svelte-comments textarea { box-sizing: border-box; width: 100%; resize: vertical; padding: 6px 7px; border: 1px solid var(--pptx-border, var(--office-border, #d1d5db)); border-radius: var(--pptx-radius, var(--office-radius, 4px)); background: var(--pptx-muted, var(--office-surface, #f3f4f6)); color: inherit; font: inherit; font-size: 12px; }
	.pptx-svelte-comments-compose button { justify-self: end; padding: 4px 8px; border: 0; border-radius: var(--pptx-radius, var(--office-radius, 4px)); background: var(--pptx-primary, var(--office-accent, #2563eb)); color: var(--pptx-primary-foreground, var(--office-accent-foreground, #ffffff)); cursor: pointer; font: inherit; font-size: 11px; font-weight: 600; }
	.pptx-svelte-comments-compose button:disabled { cursor: default; opacity: .45; }
	.pptx-svelte-comments-list { display: grid; gap: 5px; max-height: 230px; overflow-y: auto; }
	.pptx-svelte-comment-card { display: grid; gap: 5px; padding: 7px; border: 1px solid var(--pptx-border, var(--office-border, #d1d5db)); border-radius: var(--pptx-radius, var(--office-radius, 4px)); background: var(--pptx-muted, var(--office-surface, #f3f4f6)); }
	.pptx-svelte-comment-card.resolved { opacity: .67; }
	.pptx-svelte-comment-meta, .pptx-svelte-comment-actions { display: flex; align-items: center; gap: 6px; }
	.pptx-svelte-comment-meta { justify-content: space-between; font-size: 11px; }
	.pptx-svelte-comment-meta span { color: #9be9a8; font-size: 10px; font-weight: 600; }
	.pptx-svelte-comment-card p { margin: 0; font-size: 11.5px; line-height: 1.35; white-space: pre-wrap; }
	.pptx-svelte-comment-actions button { padding: 0; border: 0; background: transparent; color: var(--pptx-primary, var(--office-accent, #2563eb)); cursor: pointer; font: inherit; font-size: 10.5px; }
	.pptx-svelte-comment-actions button:last-child { color: #fb7185; }
	.pptx-svelte-comment-reply-compose { display: grid; gap: 5px; }
	.pptx-svelte-comment-reply-buttons { display: flex; justify-content: flex-end; gap: 6px; }
	.pptx-svelte-comment-reply-cancel { padding: 3px 7px; border: 1px solid var(--pptx-border, var(--office-border, #d1d5db)); border-radius: var(--pptx-radius, var(--office-radius, 4px)); background: transparent; color: inherit; cursor: pointer; font: inherit; font-size: 10.5px; }
	.pptx-svelte-comment-reply-submit { padding: 3px 8px; border: 0; border-radius: var(--pptx-radius, var(--office-radius, 4px)); background: var(--pptx-primary, var(--office-accent, #2563eb)); color: var(--pptx-primary-foreground, var(--office-accent-foreground, #ffffff)); cursor: pointer; font: inherit; font-size: 10.5px; font-weight: 600; }
	.pptx-svelte-comment-reply-submit:disabled { cursor: default; opacity: .45; }

	/*
	 * Touch target: matches MIN_TOUCH_TARGET_PX (44) from pptx-viewer-shared's
	 * render/responsive module. `!important` is load-bearing: see
	 * ViewerGlobalStyles.svelte's document-level baseline reset note
	 * elsewhere in this codebase (out-specificities a plain scoped selector).
	 */
	@media (max-width: 767px) {
		.pptx-svelte-comments-compose button,
		.pptx-svelte-comment-actions button,
		.pptx-svelte-comment-reply-cancel,
		.pptx-svelte-comment-reply-submit {
			min-width: 44px !important;
			min-height: 44px !important;
		}
	}
</style>
