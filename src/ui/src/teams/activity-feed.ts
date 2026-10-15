import { html, nothing } from 'lit';
import {
	ACTIVITY_FILTERS,
	type ActivityFilter,
	type ActivityItem,
	type ActivityKind,
	describeActivity,
	filterActivity,
} from 'ooxml-core/teams';
import { definer } from '../registry';
import { TeamsElement, dayLabel, formatTime, withStyles } from './base';
import { icon } from './icons';
import css from './activity-feed.css?raw';

const FILTER_LABEL: Record<ActivityFilter, string> = {
	all: 'All',
	unread: 'Unread',
	mentions: 'Mentions',
	replies: 'Replies',
	reactions: 'Reactions',
};
const KIND_ICON: Record<ActivityKind, string> = {
	mention: 'at',
	reply: 'reply',
	reaction: 'smile',
	pin: 'pin',
};
const EMPTY: Record<ActivityFilter, string> = {
	all: 'No activity yet. Mentions, replies, reactions and pinned messages show up here.',
	unread: 'You are all caught up.',
	mentions: 'Nobody has mentioned you yet.',
	replies: 'No replies to your messages or followed threads yet.',
	reactions: 'No reactions to your messages yet.',
};

/** "14:05" today, "Yesterday" or a short date otherwise. */
export function activityTime(ts: number, now = Date.now()): string {
	const day = dayLabel(ts, now);
	if (day === 'Today') return formatTime(ts);
	if (day === 'Yesterday') return day;
	return new Date(ts).toLocaleDateString([], { month: 'short', day: 'numeric' });
}

/**
 * The Teams Activity feed: filter pills (All, Unread, Mentions, Replies, Reactions), a list of
 * what happened to you, newest first, and Mark all as read. Properties: `items` (from
 * `TeamsState.activity`), `filter`, `selectedId`. Events (bubbling, composed):
 * `office-activity-open` `{ id }`, `office-activity-read` `{ id, read }` (the per-row toggle),
 * `office-activity-read-all` `{}` and `office-activity-filter` `{ filter }`.
 * Message text is rendered as text only.
 */
export class OfficeUiActivityFeed extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		items: { attribute: false },
		filter: { type: String, reflect: true },
		selectedId: { type: String, attribute: 'selected-id' },
	};
	declare items: ActivityItem[];
	declare filter: ActivityFilter;
	declare selectedId: string;

	constructor() {
		super();
		this.items = [];
		this.filter = 'all';
		this.selectedId = '';
	}

	private get unread(): number {
		return this.items.filter((item) => !item.read).length;
	}

	protected override render() {
		const shown = filterActivity(this.items, this.filter);
		const unread = this.unread;
		return html`
			<header>
				<h2>Activity</h2>
				<button
					type="button"
					class="read-all"
					?disabled=${unread === 0}
					@click=${() => this.fire('office-activity-read-all', {})}
				>
					Mark all as read
				</button>
			</header>
			<div class="filters" role="group" aria-label="Filter activity">
				${ACTIVITY_FILTERS.map(
					(value) =>
						html`<button
							type="button"
							aria-pressed=${String(this.filter === value)}
							@click=${() => {
								this.filter = value;
								this.fire('office-activity-filter', { filter: value });
							}}
						>
							${FILTER_LABEL[value]}${value === 'unread' && unread ? html` <span class="count">${unread}</span>` : nothing}
						</button>`,
				)}
			</div>
			${
				shown.length
					? html`<ul aria-label="Activity feed">
							${shown.map((item) => this.row(item))}
						</ul>`
					: html`<p class="empty" role="status">${EMPTY[this.filter]}</p>`
			}
		`;
	}

	private row(item: ActivityItem) {
		const line = describeActivity(item);
		return html`<li data-unread=${String(!item.read)} data-kind=${item.kind}>
			<button
				type="button"
				class="item"
				aria-current=${this.selectedId === item.id ? 'true' : 'false'}
				aria-label=${`${item.read ? '' : 'Unread. '}${line}. ${item.preview}`}
				@click=${() => this.fire('office-activity-open', { id: item.id })}
			>
				<span class="face">
					<office-ui-avatar
						size="md"
						name=${item.actorName}
						seed=${item.actorId}
					></office-ui-avatar>
					<span class="kind" aria-hidden="true">${icon(KIND_ICON[item.kind])}</span>
				</span>
				<span class="text">
					<span class="line">${line}</span>
					<span class="preview">${item.preview}</span>
				</span>
				<span class="when">
					<time datetime=${new Date(item.ts).toISOString()}>${activityTime(item.ts)}</time>
					${item.read ? nothing : html`<span class="dot" aria-hidden="true"></span>`}
				</span>
			</button>
			<button
				type="button"
				class="toggle"
				title=${item.read ? 'Mark as unread' : 'Mark as read'}
				aria-label=${item.read ? 'Mark as unread' : 'Mark as read'}
				@click=${() => this.fire('office-activity-read', { id: item.id, read: !item.read })}
			>
				${icon(item.read ? 'bell' : 'check')}
			</button>
		</li>`;
	}
}

export const defineActivityFeed = definer('office-ui-activity-feed', () => OfficeUiActivityFeed);
