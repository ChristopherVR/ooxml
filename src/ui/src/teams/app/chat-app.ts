// Template helpers of the Chat app, kept out of teams-app.ts: the conversation header of a 1:1 or
// group chat and the meeting cards a chat list shows. Pure functions of core state.
import { html, nothing } from 'lit';
import { type ChatView, type MeetingView, occurrencesBetween } from 'ooxml-core/teams';
import type { MeetingCard } from '../chat-list';
import { icon } from '../icons';

export interface ChatHeaderHandlers {
	back: () => void;
	call: () => void;
	people: () => void;
}

/** The header above a chat's messages: who it is with, presence, Call and People. */
export function chatHeader(chat: ChatView, on: ChatHeaderHandlers) {
	const other = chat.kind === 'direct' ? chat.members.find((m) => !m.self) : undefined;
	const online = chat.members.filter((m) => !m.self && m.availability !== 'offline').length;
	return html`<header class="channel-head chat-head">
		<button
			type="button"
			class="back"
			aria-label="Back to chats"
			title="Back to chats"
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
		<span class="topic"
			>${
				chat.kind === 'direct'
					? (other?.availability ?? '')
					: `${chat.members.length} people${online ? `, ${online} online` : ''}`
			}</span
		>
		<div class="people" aria-label="People in this chat">
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
		<button type="button" class="chat-people" @click=${on.people}>${icon('users')} People</button>
		<button type="button" class="meet" @click=${on.call}>
			${chat.live ? 'Join call' : 'Call'}
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
