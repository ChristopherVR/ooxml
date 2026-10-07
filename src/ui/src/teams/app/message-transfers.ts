import { html } from 'lit';
import type { TeamsClient, TeamsState } from 'ooxml-core/teams';

export function messageTransfers(state: TeamsState, client: TeamsClient) {
	return state.messageTransfers
		.filter((transfer) => transfer.context.channelId === state.selectedChannelId)
		.map(
			(transfer) =>
				html`<div class="message-transfer" role="status">
					<span
						>${`${transfer.context.threadId || transfer.context.replyTo ? 'Sending thread reply' : 'Sending channel post'}: ${transfer.progress.fileName}. ${transfer.progress.completed} of ${transfer.progress.total} files uploaded`}</span
					><progress
						aria-label="Message file progress"
						max=${transfer.progress.total}
						value=${transfer.progress.completed}
					></progress
					><button type="button" @click=${() => client.cancelSend(transfer.id)}>
						Cancel message send
					</button>
				</div>`,
		);
}
