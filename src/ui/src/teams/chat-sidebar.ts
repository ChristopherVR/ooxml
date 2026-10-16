import { html, nothing } from 'lit';
import type { ChatView } from 'ooxml-core/teams';
import { definer } from '../registry';
import { TeamsElement, dayKind, formatTime, withStyles } from './base';
import { teamsI18n } from './i18n/translator';
import { icon } from './icons';
import css from './chat-sidebar.css?raw';

/** The time a row shows: the clock today, otherwise a short date. */
function stamp(ts: number, locale: string): string {
	if (!ts) return '';
	return dayKind(ts) === 'today'
		? formatTime(ts, locale)
		: new Date(ts).toLocaleDateString(teamsI18n(locale).intl, { month: 'short', day: 'numeric' });
}

/**
 * The Chat app's list: 1:1 and group chats, most recent first, with presence on the avatar, the
 * last message and an unread badge. Properties: `chats` (core `ChatView[]`), `selectedId`.
 * Emits `office-chat-select` `{ id }` and `office-chat-new` `{}` (the host asks who to chat with).
 */
export class OfficeUiChatSidebar extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		chats: { attribute: false },
		selectedId: { type: String, attribute: 'selected-id' },
	};
	declare chats: ChatView[];
	declare selectedId: string;

	constructor() {
		super();
		this.chats = [];
		this.selectedId = '';
	}

	protected override render() {
		return html`
			<header>
				<h2>${this.t('rail.chat')}</h2>
				<button
					class="new"
					type="button"
					title=${this.t('chat.new')}
					aria-label=${this.t('chat.new')}
					@click=${() => this.fire('office-chat-new', {})}
				>
					${icon('plus')}
				</button>
			</header>
			${
				this.chats.length === 0
					? html`<p class="empty">${this.t('chat.none')}</p>`
					: html`<ul>
							${this.chats.map((c) => this.row(c))}
						</ul>`
			}
		`;
	}

	private row(c: ChatView) {
		const other = c.kind === 'direct' ? c.members.find((m) => !m.self) : undefined;
		const preview = c.lastMessage
			? c.lastMessage.system
				? c.lastMessage.text
				: c.lastMessage.mine
					? this.t('chat.previewYou', { text: c.lastMessage.text })
					: c.kind === 'direct'
						? c.lastMessage.text
						: this.t('chat.previewAuthor', {
								author: c.lastMessage.authorName,
								text: c.lastMessage.text,
							})
			: this.t('chat.noMessagesPreview');
		return html`
			<li>
				<button
					class="row"
					type="button"
					aria-current=${c.id === this.selectedId ? 'true' : 'false'}
					data-unread=${c.unread ? 'true' : 'false'}
					@click=${() => this.fire('office-chat-select', { id: c.id })}
				>
					${
						other
							? html`<office-ui-avatar
									name=${other.name}
									seed=${other.id}
									presence=${other.availability}
								></office-ui-avatar>`
							: html`<span class="group" aria-hidden="true"
									>${icon(c.meetingId ? 'calendar' : 'users')}</span
								>`
					}
					<span class="text">
						<span class="line">
							<span class="title">${c.title}</span>
							<time>${stamp(c.ts, this.locale)}</time>
						</span>
						<span class="preview">${preview}</span>
					</span>
					${
						c.live
							? html`<span class="live" title=${this.t('call.inProgress')}>${icon('video')}</span>`
							: nothing
					}
					${
						c.unread
							? html`<span
									class="count"
									aria-label=${this.t('chat.unreadCount', { count: c.unread })}
									>${c.unread > 99 ? '99+' : c.unread}</span
								>`
							: nothing
					}
				</button>
			</li>
		`;
	}
}

export const defineChatSidebar = definer('office-ui-chat-sidebar', () => OfficeUiChatSidebar);
