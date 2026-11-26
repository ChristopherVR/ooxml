<script setup lang="ts">
import { attachCommentsTaskPane, createCommentsPanelController } from 'ooxml-ui/pptx';
import type { PptxComment, PptxCommentMention, PptxModernCommentAuthor } from 'pptx-viewer-core';
import { formatCommentTimestamp } from 'pptx-viewer-core';
import { onMounted, onBeforeUnmount, ref, computed, inject } from 'vue';
import { useI18n } from 'vue-i18n';

import { COMMENTS_PANEL_CONTEXT } from '../composables/comments-panel-context';
import { useController } from '../composables/useController';
import CommentBody from './CommentBody.vue';
import CommentMentionTextarea from './CommentMentionTextarea.vue';

/**
 * CommentsPanel: side panel listing the active slide's comments.
 *
 * It renders the supplied `comments` (the active slide's, sliced by the host); the drafts, the
 * comment being edited, the open reply composer and every command are the viewer's shared
 * comments panel controller, provided by `useCommentsWiring`, so a half-typed comment stays with
 * its slide and every panel under the viewer shows the same drafts.
 */
const props = defineProps<{
	comments: PptxComment[];
	authorName: string;
	/**
	 * True when hosted inside a chrome that already renders its own title +
	 * close button (the mobile `MobileSheet`). Suppresses the internal
	 * `<header>` so mobile doesn't show "Comments" twice stacked.
	 */
	embedded?: boolean;
	closeTaskPane?: () => void;
	/** Modern comment authors (`ppt/commentAuthors.xml`), for the `@`-mention typeahead. */
	modernCommentAuthors?: PptxModernCommentAuthor[];
}>();

const { t } = useI18n();

// Outside a viewer (a host mounting the panel on its own) there is no deck to write to, so the
// panel lists the comments read-only.
const context = inject(COMMENTS_PANEL_CONTEXT, null) ?? {
	controller: createCommentsPanelController({
		slides: () => [],
		canEdit: () => false,
		authorName: () => props.authorName,
		commit: () => undefined,
	}),
	slideIndex: () => 0,
	slideId: () => undefined,
	canEdit: () => false,
};
const { controller } = context;
const state = useController(controller);
const slideId = computed(() => context.slideId() ?? '');
const slideIndex = (): number => context.slideIndex();
const canEdit = computed(() => context.canEdit());

const mentionAuthors = computed(() => props.modernCommentAuthors ?? []);

const draft = computed({
	get: () => state.value.draftBySlideId[slideId.value] ?? '',
	set: (text: string) => controller.setDraft(slideId.value, text),
});
const draftMentions = computed(
	() => (state.value.draftMentionsBySlideId[slideId.value] ?? []) as PptxCommentMention[],
);
const setDraftMentions = (mentions: PptxCommentMention[]): void =>
	controller.setDraft(slideId.value, draft.value, mentions);
const canAdd = computed<boolean>(() => canEdit.value && draft.value.trim().length > 0);
const submit = (): void => {
	controller.add(slideIndex());
};

const editingId = computed(() => state.value.editingIdBySlideId[slideId.value] ?? null);
const editDraft = (id: string): string => state.value.editDraftById[id] ?? '';

const replyingTo = computed(() => state.value.replyingToId);
const replyDraft = (id: string): string => state.value.replyDraftById[id] ?? '';
const replyMentions = (id: string): PptxCommentMention[] =>
	(state.value.replyMentionsById[id] ?? []) as PptxCommentMention[];

const formatTimestamp = (value: string | undefined): string => formatCommentTimestamp(value);
const taskPane = ref<HTMLElement | null>(null);
let detachTaskPane: (() => void) | undefined;
onMounted(() => {
	if (taskPane.value && props.closeTaskPane) {
		detachTaskPane = attachCommentsTaskPane(taskPane.value, {
			label: t('pptx.toolbar.comments'),
			closeLabel: t('pptx.common.close'),
			close: () => props.closeTaskPane?.(),
		});
	}
});
onBeforeUnmount(() => detachTaskPane?.());
</script>

<template>
	<aside
		ref="taskPane"
		class="pptx-comments-panel flex h-full min-h-0 w-full flex-col border-l border-border bg-card text-foreground"
		:aria-label="t('pptx.comments.slideComments')"
	>
		<header
			v-if="!embedded"
			class="pptx-comments-panel__header flex items-center justify-between border-b border-border px-4 py-3"
		>
			<h2 class="pptx-comments-panel__title m-0 text-sm font-semibold">
				{{ t('pptx.toolbar.comments') }}
			</h2>
			<span
				class="pptx-comments-panel__count text-xs text-muted-foreground"
				data-testid="comment-count"
				:class="{ 'mr-8': closeTaskPane }"
			>
				{{ props.comments.length }}
			</span>
		</header>

		<ul
			v-if="props.comments.length > 0"
			class="pptx-comments-panel__list m-0 min-h-0 flex-1 list-none overflow-y-auto p-2"
		>
			<li
				v-for="comment in props.comments"
				:key="comment.id"
				class="pptx-comments-panel__item mb-2 rounded-lg border border-border px-3 py-2.5"
				:class="{ 'pptx-comments-panel__item--resolved opacity-60': comment.resolved }"
				:data-comment-id="comment.id"
			>
				<div class="pptx-comments-panel__meta mb-1 flex items-baseline justify-between gap-2">
					<span class="pptx-comments-panel__author text-[13px] font-semibold">{{
						comment.author || t('pptx.comments.unknownAuthor')
					}}</span>
					<time
						v-if="formatTimestamp(comment.createdAt)"
						class="pptx-comments-panel__time text-[11px] text-muted-foreground"
					>
						{{ formatTimestamp(comment.createdAt) }}
					</time>
				</div>
				<div v-if="editingId === comment.id" class="mb-2 flex flex-col gap-1.5">
					<textarea
						class="pptx-comments-panel__edit w-full resize-y rounded-md border border-border bg-background p-2 text-[13px] text-foreground"
						rows="2"
						:aria-label="t('pptx.comments.edit')"
						:value="editDraft(comment.id)"
						@input="
							controller.setEditDraft(comment.id, ($event.target as HTMLTextAreaElement).value)
						"
					></textarea>
					<div class="flex justify-end gap-2">
						<button
							type="button"
							class="cursor-pointer rounded-md border border-border bg-transparent px-2 py-1 text-xs text-foreground hover:bg-muted"
							@click="controller.cancelEdit(slideId)"
						>
							{{ t('pptx.comments.cancel') }}
						</button>
						<button
							type="button"
							class="cursor-pointer rounded-md border-none bg-primary px-2.5 py-1 text-xs text-primary-foreground disabled:cursor-not-allowed disabled:opacity-50"
							:disabled="editDraft(comment.id).trim().length === 0"
							@click="controller.saveEdit(slideIndex(), comment.id)"
						>
							{{ t('pptx.comments.save') }}
						</button>
					</div>
				</div>
				<p
					v-else
					class="pptx-comments-panel__text m-0 mb-2 whitespace-pre-wrap break-words text-[13px]"
				>
					<CommentBody :text="comment.text" :mentions="comment.mentions" />
				</p>

				<!-- Threaded replies -->
				<ul
					v-if="comment.replies && comment.replies.length > 0"
					class="pptx-comments-panel__replies m-0 mb-2 list-none border-l-2 border-border pl-3"
				>
					<li
						v-for="reply in comment.replies"
						:key="reply.id"
						class="pptx-comments-panel__reply mb-1.5"
						:data-comment-id="reply.id"
					>
						<div class="mb-0.5 flex items-baseline justify-between gap-2">
							<span class="text-[12px] font-semibold">{{
								reply.author || t('pptx.comments.unknownAuthor')
							}}</span>
							<time
								v-if="formatTimestamp(reply.createdAt)"
								class="text-[11px] text-muted-foreground"
							>
								{{ formatTimestamp(reply.createdAt) }}
							</time>
						</div>
						<p class="m-0 whitespace-pre-wrap break-words text-[12px]">
							<CommentBody :text="reply.text" :mentions="reply.mentions" />
						</p>
					</li>
				</ul>

				<div class="pptx-comments-panel__actions flex gap-2">
					<button
						type="button"
						class="pptx-comments-panel__action cursor-pointer rounded-md border border-border bg-transparent px-2 py-1 text-xs text-foreground hover:bg-muted max-md:min-h-[44px]!"
						:data-comment-id="comment.id"
						:aria-pressed="comment.resolved ? 'true' : 'false'"
						@click="controller.toggleResolved(slideIndex(), comment.id)"
					>
						{{ comment.resolved ? t('pptx.comments.reopen') : t('pptx.comments.resolve') }}
					</button>
					<button
						type="button"
						class="pptx-comments-panel__action cursor-pointer rounded-md border border-border bg-transparent px-2 py-1 text-xs text-foreground hover:bg-muted max-md:min-h-[44px]!"
						:data-comment-id="comment.id"
						@click="controller.startReply(comment.id)"
					>
						{{ t('pptx.comments.reply') }}
					</button>
					<button
						v-if="canEdit && editingId !== comment.id"
						type="button"
						class="pptx-comments-panel__action pptx-comments-panel__action--edit cursor-pointer rounded-md border border-border bg-transparent px-2 py-1 text-xs text-foreground hover:bg-muted max-md:min-h-[44px]!"
						:data-comment-id="comment.id"
						@click="controller.startEdit(slideId, comment.id)"
					>
						{{ t('pptx.comments.edit') }}
					</button>
					<button
						type="button"
						class="pptx-comments-panel__action pptx-comments-panel__action--danger cursor-pointer rounded-md border border-border bg-transparent px-2 py-1 text-xs text-red-400 hover:bg-muted max-md:min-h-[44px]!"
						:data-comment-id="comment.id"
						:aria-label="t('pptx.comments.removeComment')"
						@click="controller.remove(slideIndex(), comment.id)"
					>
						{{ t('pptx.comments.remove') }}
					</button>
				</div>

				<!-- Reply composer -->
				<div v-if="replyingTo === comment.id" class="mt-2 flex flex-col gap-1.5">
					<CommentMentionTextarea
						:model-value="replyDraft(comment.id)"
						:mentions="replyMentions(comment.id)"
						:authors="mentionAuthors"
						textarea-class="w-full resize-y rounded-md border border-border bg-background p-2 text-[12px] text-foreground"
						:rows="2"
						:placeholder="t('pptx.comments.replyPlaceholder')"
						:aria-label="t('pptx.comments.reply')"
						@update:model-value="controller.setReplyDraft(comment.id, $event)"
						@update:mentions="controller.setReplyDraft(comment.id, replyDraft(comment.id), $event)"
					/>
					<div class="flex justify-end gap-2">
						<button
							type="button"
							class="cursor-pointer rounded-md border border-border bg-transparent px-2 py-1 text-xs text-foreground hover:bg-muted"
							@click="controller.cancelReply()"
						>
							{{ t('pptx.comments.cancel') }}
						</button>
						<button
							type="button"
							class="cursor-pointer rounded-md border-none bg-primary px-2.5 py-1 text-xs text-primary-foreground disabled:cursor-not-allowed disabled:opacity-50"
							:disabled="replyDraft(comment.id).trim().length === 0"
							@click="controller.submitReply(slideIndex(), comment.id)"
						>
							{{ t('pptx.comments.reply') }}
						</button>
					</div>
				</div>
			</li>
		</ul>

		<p
			v-else
			class="pptx-comments-panel__empty flex-1 p-4 text-[13px] text-muted-foreground"
			data-testid="comments-empty"
		>
			{{ t('pptx.comments.noneOnSlide') }}
		</p>

		<form
			data-pptx-comment-draft
			class="pptx-comments-panel__compose flex flex-col gap-2 border-t border-border px-4 py-3"
			@submit.prevent="submit"
		>
			<label
				class="pptx-comments-panel__compose-label text-xs font-semibold"
				:title="t('pptx.comments.commentingAs', { name: props.authorName })"
			>
				{{ t('pptx.comments.addComment') }}
			</label>
			<CommentMentionTextarea
				v-model="draft"
				:mentions="draftMentions"
				:authors="mentionAuthors"
				textarea-class="pptx-comments-panel__textarea w-full resize-y rounded-md border border-border bg-background p-2 text-[13px] text-foreground"
				:rows="3"
				:placeholder="t('pptx.comments.addCommentPlaceholder')"
				:aria-label="t('pptx.comments.addComment')"
				@update:mentions="setDraftMentions"
			/>
			<button
				type="submit"
				class="pptx-comments-panel__submit cursor-pointer self-end rounded-md border-none bg-primary px-3.5 py-1.5 text-[13px] text-primary-foreground disabled:cursor-not-allowed disabled:opacity-50 max-md:min-h-[44px]!"
				:disabled="!canAdd"
				data-testid="add-comment"
			>
				{{ t('pptx.comments.addComment') }}
			</button>
		</form>
	</aside>
</template>
