// Template helpers of the Chat app, kept out of teams-app.ts: the conversation header of a 1:1 or
// group chat and the meeting cards a chat list shows. Pure functions of core state.
import { html, nothing } from 'lit';
import { type ChatView, type MeetingView, occurrencesBetween } from 'ooxml-core/teams';
import type { MeetingCard } from '../chat-list';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';
import { icon } from '../icons';

export interface ChatHeaderHandlers {
	back: () => void;
	call: () => void;
	people: () => void;
}

const PRESENCE_KEY = {
	available: 'presence.available',
	busy: 'presence.busy',
	dnd: 'presence.dnd',
	brb: 'presence.brb',
	away: 'presence.away',
	offline: 'presence.offline',
} as const;

/** The header above a chat's messages: who it is with, presence, Call and People. */
export function chatHeader(chat: ChatView, on: ChatHeaderHandlers, i18n: TeamsI18n = teamsI18n()) {
	const { t } = i18n;
	const other = chat.kind === 'direct' ? chat.members.find((m) => !m.self) : undefined;
	const online = chat.members.filter((m) => !m.self && m.availability !== 'offline').length;
	const topic =
		chat.kind === 'direct'
			? other
				? t(other.activity === 'meeting' ? 'presence.inMeeting' : PRESENCE_KEY[other.availability])
				: ''
			: online
				? t('chat.peopleOnline', { count: chat.members.length, online })
				: i18n.tn('chat.peopleCount', chat.members.length);
	return html`<header class="channel-head chat-head">
		<button
			type="button"
			class="back"
			aria-label=${t('chat.backToChats')}
			title=${t('chat.backToChats')}
			@click=${on.back}
		>
			${icon('back')}
		</button>
		${
			other
				? html`<office-ui-avatar
						name=${other.name}
						seed=${other.id}
						presence=${other.availability}
					></office-ui-avatar>`
				: html`<span class="group-mark" aria-hidden="true"
						>${icon(chat.meetingId ? 'calendar' : 'users')}</span
					>`
		}
		<h1>${chat.title}</h1>
		<span class="topic">${topic}</span>
		<div class="people" aria-label=${t('chat.peopleInChat')}>
			${chat.members
				.slice(0, 5)
				.map(
					(m) =>
						html`<office-ui-avatar
							size="sm"
							name=${m.name}
							seed=${m.id}
							presence=${m.availability}
						></office-ui-avatar>`,
				)}
			${chat.members.length > 5 ? html`<span class="more">+${chat.members.length - 5}</span>` : nothing}
		</div>
		<button type="button" class="chat-people" @click=${on.people}>
			${icon('users')} ${t('chat.people')}
		</button>
		<button type="button" class="meet" @click=${on.call}>
			${t(chat.live ? 'call.joinCall' : 'chat.call')}
		</button>
	</header>`;
}

/** Meeting cards for the chat list: each meeting's next occurrence (or its start once over). */
export function meetingCards(
	meetings: readonly MeetingView[],
	now: number,
): Record<string, MeetingCard> {
	const out: Record<string, MeetingCard> = {};
	for (const m of meetings) {
		const next = occurrencesBetween(m, now, now + 400 * 86_400_000)[0];
		out[m.id] = {
			title: m.title,
			start: next?.start ?? m.start,
			end: next?.end ?? m.end,
			live: m.live,
		};
	}
	return out;
}
