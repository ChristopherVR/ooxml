import { html, nothing } from 'lit';
import type { Attachment, TeamsClient, TeamsState } from 'ooxml-core/teams';

/** A thread shares the existing message controls and core actions. No document mapping lives here. */
export function threadPane(
	state: TeamsState,
	client: TeamsClient,
	openFile: (attachment: Attachment) => void,
	close: () => void,
) {
	const thread = state.thread;
	if (!thread) return nothing;
	return html`<aside class="thread-pane" aria-label="Thread">
		<header class="thread-head">
			<h2 tabindex="-1" data-thread-heading>Thread</h2>
			<button
				type="button"
				@click=${() => client.followThread(state.selectedChannelId, thread.root.id, !state.threadFollowed)}
			>
				${state.threadFollowed ? 'Unfollow thread' : 'Follow thread'}
			</button>
			<button type="button" @click=${close}>Close thread</button>
		</header>
		<office-ui-chat-list
			.messages=${[thread.root, ...thread.replies]}
			self-id=${state.user.id}
			@office-chat-react=${(e: CustomEvent<{ messageId: string; emoji: string }>) => client.toggleReaction(e.detail.messageId, e.detail.emoji)}
			@office-chat-reply=${(e: CustomEvent<{ messageId: string }>) => client.startReply(e.detail.messageId)}
			@office-chat-edit=${(e: CustomEvent<{ messageId: string }>) => client.startEdit(e.detail.messageId)}
			@office-chat-delete=${(e: CustomEvent<{ messageId: string }>) => globalThis.confirm?.('Delete this message?') !== false && client.deleteMessage(e.detail.messageId)}
			@office-chat-open-file=${(e: CustomEvent<{ attachment: Attachment }>) => openFile(e.detail.attachment)}
		></office-ui-chat-list>
		<office-ui-chat-composer
			.typing=${state.typing}
			.replyingTo=${state.replyingTo?.authorName ?? null}
			.editing=${state.editing !== null}
			.value=${state.editing?.text ?? ''}
			placeholder="Reply in thread"
			@office-chat-typing=${() => client.notifyTyping()}
			@office-chat-cancel=${() => client.cancelCompose()}
			@office-chat-send=${(e: CustomEvent<{ text: string; files: File[] }>) => {
				const current = client.getState();
				if (!current.editing && !current.replyingTo) client.startReply(thread.root.id);
				void client.send(e.detail);
			}}
		></office-ui-chat-composer>
	</aside>`;
}
