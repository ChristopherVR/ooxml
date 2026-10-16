import { html } from 'lit';
import type { DraftContext, TeamsClient, TeamsState } from 'ooxml-core/teams';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';

export function draftList(
	state: TeamsState,
	client: TeamsClient,
	open: (context: DraftContext) => void,
	i18n: TeamsI18n = teamsI18n(),
) {
	const { t, tn } = i18n;
	return html`<section class="followed-threads" aria-label=${t('drafts.title')}>
		<h1>${t('drafts.title')}</h1>
		<p>${t('drafts.intro')}</p>
		<ul>
			${
				state.drafts.length
					? state.drafts.map((draft) => {
							const channel = state.channels.find((item) => item.id === draft.context.channelId)!;
							const body = draft.text || t('drafts.attachments');
							return html`<li>
								<button
									type="button"
									aria-label=${t('drafts.resume', {
										channel: channel.name,
										text: draft.text.slice(0, 80) || t('drafts.attachments'),
									})}
									@click=${() => open(draft.context)}
								>
									<strong># ${channel.name}</strong><span>${body}</span
									><small
										>${draft.context.editId ? t('drafts.edit') : draft.context.threadId ? t('drafts.threadReply') : t('drafts.newPost')}
										·
										${tn('drafts.attachmentCount', draft.files.length + draft.missingFiles.length)}</small
									></button
								><button
									type="button"
									@click=${() => {
										if (globalThis.confirm?.(t('drafts.confirmDiscard')) !== false)
											client.discardDraft(draft.context);
									}}
								>
									${t('drafts.discard')}
								</button>
							</li>`;
						})
					: html`<li>${t('drafts.none')}</li>`
			}
		</ul>
	</section>`;
}
