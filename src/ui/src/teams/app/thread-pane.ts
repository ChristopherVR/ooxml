import { html, nothing } from 'lit';
import type { Attachment, ChatDraft, TeamsClient, TeamsState } from 'ooxml-core/teams';
import type { ChatSendDetail } from '../chat-composer';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';

/** A thread shares the existing message controls and core actions. No document mapping lives here. */
export function threadPane(
	state: TeamsState,
	client: TeamsClient,
	openFile: (attachment: Attachment) => void,
	close: () => void,
	highlightId = '',
	i18n: TeamsI18n = teamsI18n(),
) {
	const thread = state.thread;
	if (!thread) return nothing;
	const { t } = i18n;
	return html`<aside class="thread-pane" aria-label=${t('thread.title')}>
		<header class="thread-head">
			<h2 tabindex="-1" data-thread-heading>${t('thread.title')}</h2>
			<button
				type="button"
				@click=${() => client.followThread(state.selectedChannelId, thread.root.id, !state.threadFollowed)}
			>
				${t(state.threadFollowed ? 'thread.unfollow' : 'thread.follow')}
			</button>
			<button type="button" @click=${close}>${t('thread.close')}</button>
		</header>
		<office-ui-chat-list
			.messages=${[thread.root, ...thread.replies]}
			.pinnedIds=${state.pinned.map((p) => p.message.id)}
			self-id=${state.user.id}
			.highlightId=${highlightId}
			@office-chat-pin=${(e: CustomEvent<{ messageId: string }>) => client.pinMessage(e.detail.messageId)}
			@office-chat-unpin=${(e: CustomEvent<{ messageId: string }>) => client.unpinMessage(e.detail.messageId)}
			@office-chat-react=${(e: CustomEvent<{ messageId: string; emoji: string }>) => client.toggleReaction(e.detail.messageId, e.detail.emoji)}
			@office-chat-reply=${(e: CustomEvent<{ messageId: string }>) => client.startReply(e.detail.messageId)}
			@office-chat-edit=${(e: CustomEvent<{ messageId: string }>) => client.startEdit(e.detail.messageId)}
			@office-chat-delete=${(e: CustomEvent<{ messageId: string }>) => globalThis.confirm?.(t('chat.confirmDelete')) !== false && client.deleteMessage(e.detail.messageId)}
			@office-chat-open-file=${(e: CustomEvent<{ attachment: Attachment }>) => openFile(e.detail.attachment)}
		></office-ui-chat-list>
		<office-ui-chat-composer
			.typing=${state.typing}
			.replyingTo=${state.replyingTo?.authorName ?? null}
			.editing=${state.editing !== null}
			.value=${state.draft.text}
			.files=${state.draft.files}
			.missingFiles=${state.draft.missingFiles}
			.mentionTargets=${state.mentionTargets}
			@office-chat-draft=${(e: CustomEvent<ChatDraft>) => client.setDraft(e.detail)}
			placeholder=${t('thread.replyPlaceholder')}
			@office-chat-typing=${() => client.notifyTyping()}
			@office-chat-cancel=${() => client.cancelCompose()}
			@office-chat-send=${(e: CustomEvent<ChatSendDetail>) => {
				const current = client.getState();
				if (!current.editing && !current.replyingTo) client.startReply(thread.root.id);
				void client.send(e.detail);
			}}
		></office-ui-chat-composer>
	</aside>`;
}
