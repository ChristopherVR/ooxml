import { html } from 'lit';
import type { FollowedThread, TeamsClient, TeamsState } from 'ooxml-core/teams';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';

export function followedThreads(
	state: TeamsState,
	client: TeamsClient,
	open: (thread: FollowedThread) => void,
	unreadOnly: boolean,
	filter: (unreadOnly: boolean) => void,
	i18n: TeamsI18n = teamsI18n(),
) {
	const { t, tn } = i18n;
	const threads = state.followedThreads.filter((thread) => !unreadOnly || thread.unread > 0);
	return html`<section class="followed-threads" aria-label=${t('followed.title')}>
		<h1>${t('followed.title')}</h1>
		<p>${t('followed.intro')}</p>
		<fieldset>
			<legend>${t('followed.autoFollow')}</legend>
			<label
				><input
					type="checkbox"
					.checked=${state.threadFollowSettings.started}
					@change=${(event: Event) => client.setThreadFollowSettings({ started: (event.target as HTMLInputElement).checked })}
				/>
				${t('followed.started')}</label
			>
			<label
				><input
					type="checkbox"
					.checked=${state.threadFollowSettings.replied}
					@change=${(event: Event) => client.setThreadFollowSettings({ replied: (event.target as HTMLInputElement).checked })}
				/>
				${t('followed.replied')}</label
			>
		</fieldset>
		<label
			><input
				type="checkbox"
				.checked=${unreadOnly}
				@change=${(event: Event) => filter((event.target as HTMLInputElement).checked)}
			/>
			${t('followed.unreadOnly')}</label
		>
		<ul>
			${
				threads.length
					? threads.map(
							(thread) =>
								html`<li data-unread=${String(thread.unread > 0)}>
									<button
										type="button"
										@click=${() => open(thread)}
										aria-label=${t('followed.open', {
											channel: thread.channelName,
											text: thread.root.deleted ? t('followed.deletedMessage') : thread.root.text,
										})}
									>
										<strong># ${thread.channelName}</strong
										><span
											>${thread.root.deleted ? t('chat.deleted') : thread.root.text || t('followed.sharedFile')}</span
										>
										<small
											>${tn('chat.replies', thread.replyCount)} ·
											${new Date(thread.updatedAt).toLocaleString(i18n.intl)}</small
										></button
									><span
										>${thread.unread > 0 ? t('followed.unreadCount', { count: thread.unread }) : t('followed.read')}</span
									><button
										type="button"
										@click=${() => client.markThreadRead(thread.channelId, thread.root.id, thread.unread > 0)}
									>
										${t(thread.unread > 0 ? 'activity.markRead' : 'activity.markUnread')}</button
									><button
										type="button"
										@click=${() => client.followThread(thread.channelId, thread.root.id, false)}
									>
										${t('thread.unfollow')}
									</button>
								</li>`,
						)
					: html`<li>${t(unreadOnly ? 'followed.noneUnread' : 'followed.none')}</li>`
			}
		</ul>
	</section>`;
}
