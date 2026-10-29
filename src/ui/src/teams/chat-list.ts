import { html, nothing, type TemplateResult } from 'lit';
import { type Mention, mentionSegments, mentionsUser } from 'ooxml-core/teams';
import { definer } from '../registry';
import { TeamsElement, dayLabel, formatSize, formatTime, formatWhen, withStyles } from './base';
import { icon } from './icons';
import css from './chat-list.css?raw';

export interface ChatAttachment {
	name: string;
	kind: 'docx' | 'xlsx' | 'pptx' | 'vsdx' | 'other';
	size?: number;
	url?: string;
}

export interface ChatMessage {
	id: string;
	authorId: string;
	authorName: string;
	text: string;
	ts: number;
	editedAt?: number;
	deleted?: boolean;
	/** An app notice (someone joined or left): shown as a centred line, not a bubble. */
	system?: boolean;
	/** A meeting card: the scheduled meeting this message announces. */
	meetingId?: string;
	replyTo?: string;
	tabId?: string;
	attachments?: ChatAttachment[];
	/** @mentions as ranges of `text`; messages from older clients have none. */
	mentions?: Mention[];
	reactions?: Record<string, string[]>;
}

/** What a meeting card shows; a card whose meeting is missing reads as canceled. */
export interface MeetingCard {
	title: string;
	start: number;
	end: number;
	live: boolean;
}

export const QUICK_REACTIONS = ['👍', '❤️', '😂', '😮', '😢', '🎉'] as const;
const KIND_LETTER: Record<ChatAttachment['kind'], string> = {
	docx: 'W',
	xlsx: 'X',
	pptx: 'P',
	vsdx: 'V',
	other: 'F',
};
const GROUP_MS = 5 * 60_000;

/**
 * The conversation: messages as Teams-style bubbles with day separators, quoted replies,
 * reaction pills, Office file cards and a hover toolbar. Properties: `messages`, `selfId`,
 * `pinnedIds` (messages shown with a pin and an Unpin action instead of Pin). @mentions are
 * highlighted; those that address `selfId` (or everyone) are emphasised and mark the message.
 * Events (bubbling, composed): `office-chat-react` `{ messageId, emoji }` (toggle),
 * `office-chat-reply` / `office-chat-edit` / `office-chat-delete` / `office-chat-pin` /
 * `office-chat-unpin` `{ messageId }` and
 * `office-chat-open-file` `{ attachment }`, `office-chat-open-tab` `{ tabId }`, and
 * `office-chat-meeting` `{ meetingId, action: 'join' | 'open' }` from a meeting card (`meetings`).
 * Message text is always rendered as text.
 */
export class OfficeUiChatList extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		messages: { attribute: false },
		selfId: { type: String, attribute: 'self-id' },
		replyCounts: { attribute: false },
		pinnedIds: { attribute: false },
		highlightId: { type: String, attribute: 'highlight-id' },
		meetings: { attribute: false },
		moderator: { type: Boolean },
	};
	declare messages: ChatMessage[];
	declare selfId: string;
	declare replyCounts: Record<string, number>;
	declare pinnedIds: readonly string[];
	/** A message to flash briefly, for example one opened from the Activity feed. */
	declare highlightId: string;
	/** Meeting cards by meeting id (the meetings in the person's calendar). */
	declare meetings: Record<string, MeetingCard>;
	/** The signed-in person moderates: others' messages get a Hide action (the server enforces). */
	declare moderator: boolean;
	private stickToBottom = true;

	constructor() {
		super();
		this.messages = [];
		this.selfId = '';
		this.replyCounts = {};
		this.pinnedIds = [];
		this.highlightId = '';
		this.meetings = {};
		this.moderator = false;
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.setAttribute('role', 'log');
		this.setAttribute('aria-live', 'polite');
		this.addEventListener('scroll', () => {
			this.stickToBottom = this.scrollHeight - this.scrollTop - this.clientHeight < 64;
		});
	}

	protected override hostLabel(): string {
		return this.t('chat.messages');
	}

	protected override updated(): void {
		if (this.stickToBottom) this.scrollTop = this.scrollHeight;
	}

	protected override render() {
		if (this.messages.length === 0) return html`<p class="empty">${this.t('chat.empty')}</p>`;
		const now = Date.now();
		const byId = new Map(this.messages.map((m) => [m.id, m]));
		const rows: TemplateResult[] = [];
		let previous: ChatMessage | undefined;
		for (const m of this.messages) {
			if (m.system && !m.deleted) {
				rows.push(html`<div class="system" data-message-id=${m.id} tabindex="-1">${m.text}</div>`);
				previous = undefined;
				continue;
			}
			if (!previous || dayLabel(previous.ts, now, this.locale) !== dayLabel(m.ts, now, this.locale))
				rows.push(html`<div class="day"><span>${dayLabel(m.ts, now, this.locale)}</span></div>`);
			const grouped =
				previous !== undefined &&
				previous.authorId === m.authorId &&
				m.ts - previous.ts < GROUP_MS &&
				!m.replyTo &&
				dayLabel(previous.ts, now, this.locale) === dayLabel(m.ts, now, this.locale);
			rows.push(this.message(m, grouped, byId.get(m.replyTo ?? '')));
			previous = m;
		}
		return html`${rows}`;
	}

	private message(m: ChatMessage, grouped: boolean, parent: ChatMessage | undefined) {
		const own = m.authorId === this.selfId;
		const reactions = Object.entries(m.reactions ?? {}).filter(([, users]) => users.length > 0);
		return html`
			<article
				class="msg"
				tabindex="-1"
				data-message-id=${m.id}
				data-own=${String(own)}
				data-highlight=${String(m.id === this.highlightId)}
				data-grouped=${String(grouped)}
				data-mentions-you=${String(!m.deleted && mentionsUser(m, this.selfId) && !own)}
			>
				<div class="gutter">
					${
						grouped
							? nothing
							: html`<office-ui-avatar name=${m.authorName} seed=${m.authorId}></office-ui-avatar>`
					}
				</div>
				<div class="body">
					${
						grouped
							? nothing
							: html`<header>
									<span class="author">${m.authorName}</span>
									<time>${formatTime(m.ts, this.locale)}</time>
								</header>`
					}
					${this.pinMark(m)}
					<div class="bubble">
						${
							parent
								? html`<blockquote>
										<strong>${parent.authorName}</strong
										>${parent.deleted ? this.t('chat.deleted') : parent.text.slice(0, 120)}
									</blockquote>`
								: nothing
						}
						${
							m.deleted
								? html`<span class="deleted">${this.t('chat.deleted')}</span>`
								: html`<span class="text">${this.body(m)}</span>${
											m.editedAt
												? html` <span class="edited">${this.t('chat.edited')}</span>`
												: nothing
										}`
						}
						${!m.deleted && m.meetingId ? this.meetingCard(m.meetingId) : nothing}
						${!m.deleted && m.attachments?.length ? this.files(m.attachments) : nothing}
						${!m.deleted && m.tabId ? html`<button class="thread-link" type="button" @click=${() => this.fire('office-chat-open-tab', { tabId: m.tabId })}>${this.t('chat.openTab')}</button>` : nothing}
					</div>
					${
						!m.deleted && reactions.length
							? html`<div class="reactions">
									${reactions.map(
										([emoji, users]) => html`
											<button
												type="button"
												class="pill"
												aria-pressed=${String(users.includes(this.selfId))}
												aria-label=${`${emoji} ${users.length}`}
												@click=${() => this.fire('office-chat-react', { messageId: m.id, emoji })}
											>
												<span>${emoji}</span><span>${users.length}</span>
											</button>
										`,
									)}
								</div>`
							: nothing
					}
					${m.deleted ? nothing : this.toolbar(m, own)}
					${this.replyCounts[m.id] === undefined ? nothing : html`<button type="button" class="thread-link" @click=${() => this.fire('office-chat-thread', { messageId: m.id })}>${this.tn('chat.replies', this.replyCounts[m.id]!)}</button>`}
				</div>
			</article>
		`;
	}

	/** Plain runs and highlighted mentions (the name without its `@`, as in Teams), with no whitespace between them. */
	private body(m: ChatMessage) {
		return mentionSegments(m.text, m.mentions).map((segment) =>
			segment.mention
				? html`<span
						class="mention"
						data-self=${String(segment.mention.kind !== 'user' || segment.mention.id === this.selfId)}
						>${segment.text.slice(1)}</span
					>`
				: segment.text,
		);
	}

	private meetingCard(meetingId: string) {
		const card = this.meetings[meetingId];
		if (!card)
			return html`<div class="meeting-card" data-gone="true">
				<span class="meeting-icon">${icon('calendar')}</span>
				<span class="meeting-meta">
					<span class="meeting-title">${this.t('chat.meetingGone')}</span>
				</span>
			</div>`;
		return html`<div class="meeting-card">
			<span class="meeting-icon">${icon('calendar')}</span>
			<span class="meeting-meta">
				<span class="meeting-title">${card.title}</span>
				<span class="meeting-time">${formatWhen(card.start, card.end, this.locale)}</span>
			</span>
			<button
				type="button"
				class="meeting-open"
				@click=${() => this.fire('office-chat-meeting', { meetingId, action: 'open' })}
			>
				${this.t('chat.meetingDetails')}
			</button>
			<button
				type="button"
				class="meeting-join"
				@click=${() => this.fire('office-chat-meeting', { meetingId, action: 'join' })}
			>
				${this.t(card.live ? 'call.joinNow' : 'call.join')}
			</button>
		</div>`;
	}

	private toolbar(m: ChatMessage, own: boolean) {
		return html`
			<div class="toolbar" role="toolbar" aria-label=${this.t('chat.messageActions')}>
				${QUICK_REACTIONS.slice(0, 4).map(
					(emoji) =>
						html`<button
							type="button"
							aria-label=${this.t('chat.react', { emoji })}
							@click=${() => this.fire('office-chat-react', { messageId: m.id, emoji })}
						>
							${emoji}
						</button>`,
				)}
				<button
					type="button"
					data-action="reply"
					title=${this.t('chat.reply')}
					aria-label=${this.t('chat.reply')}
					@click=${() => this.fire('office-chat-reply', { messageId: m.id })}
				>
					${icon('reply')}
				</button>
				${this.pinButton(m)}
				${
					own
						? html`<button
									type="button"
									title=${this.t('common.edit')}
									aria-label=${this.t('common.edit')}
									@click=${() => this.fire('office-chat-edit', { messageId: m.id })}
								>
									${icon('pencil')}
								</button>
								<button
									type="button"
									title=${this.t('chat.delete')}
									aria-label=${this.t('chat.delete')}
									@click=${() => this.fire('office-chat-delete', { messageId: m.id })}
								>
									${icon('trash')}
								</button>`
						: this.moderator && !m.system
							? html`<button
									type="button"
									data-action="hide"
									title=${this.t('chat.hide')}
									aria-label=${this.t('chat.hide')}
									@click=${() => this.fire('office-chat-delete', { messageId: m.id, moderation: true })}
								>
									${icon('trash')}
								</button>`
							: nothing
				}
			</div>
		`;
	}

	private pinMark(m: ChatMessage) {
		return !m.deleted && this.pinnedIds.includes(m.id)
			? html`<span class="pin-mark">${icon('pin')} ${this.t('chat.pinned')}</span>`
			: nothing;
	}

	private pinButton(m: ChatMessage) {
		const pinned = this.pinnedIds.includes(m.id);
		const label = this.t(pinned ? 'chat.unpin' : 'chat.pin');
		return html`<button
			type="button"
			title=${label}
			aria-label=${label}
			aria-pressed=${String(pinned)}
			@click=${() => this.fire(pinned ? 'office-chat-unpin' : 'office-chat-pin', { messageId: m.id })}
		>
			${icon('pin')}
		</button>`;
	}

	private files(files: ChatAttachment[]) {
		return html`<div class="files">
			${files.map(
				(a) => html`
					<button
						type="button"
						class="file"
						aria-label=${this.t('chat.openFile', { name: a.name })}
						@click=${() => this.fire('office-chat-open-file', { attachment: a })}
					>
						<span class="badge" data-kind=${a.kind}>${KIND_LETTER[a.kind]}</span>
						<span class="meta">
							<span class="file-name">${a.name}</span>
							<span class="file-size"
								>${a.kind.toUpperCase()} ${formatSize(a.size, this.locale)}</span
							>
						</span>
					</button>
				`,
			)}
		</div>`;
	}
}

export const defineChatList = definer('office-ui-chat-list', () => OfficeUiChatList);
