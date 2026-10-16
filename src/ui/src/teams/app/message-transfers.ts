import { html } from 'lit';
import type { TeamsClient, TeamsState } from 'ooxml-core/teams';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';

export function messageTransfers(
	state: TeamsState,
	client: TeamsClient,
	i18n: TeamsI18n = teamsI18n(),
) {
	const { t } = i18n;
	return state.messageTransfers
		.filter((transfer) => transfer.context.channelId === state.selectedChannelId)
		.map(
			(transfer) =>
				html`<div class="message-transfer" role="status">
					<span
						>${t('transfer.status', {
							kind: t(
								transfer.context.threadId || transfer.context.replyTo
									? 'transfer.threadReply'
									: 'transfer.channelPost',
							),
							file: transfer.progress.fileName,
							completed: transfer.progress.completed,
							total: transfer.progress.total,
						})}</span
					><progress
						aria-label=${t('transfer.progress')}
						max=${transfer.progress.total}
						value=${transfer.progress.completed}
					></progress
					><button type="button" @click=${() => client.cancelSend(transfer.id)}>
						${t('transfer.cancel')}
					</button>
				</div>`,
		);
}
