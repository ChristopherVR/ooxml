import { html, nothing } from 'lit';
import {
	ACTIVITY_FILTERS,
	type ActivityFilter,
	type ActivityItem,
	type ActivityKind,
	filterActivity,
} from 'ooxml-core/teams';
import { definer } from '../registry';
import { TeamsElement, dayKind, formatTime, withStyles } from './base';
import { describeActivityIn } from './i18n/describe';
import { teamsI18n } from './i18n/translator';
import { icon } from './icons';
import css from './activity-feed.css?raw';

const FILTER_KEY = {
	all: 'activity.filter.all',
	unread: 'activity.filter.unread',
	mentions: 'activity.filter.mentions',
	replies: 'activity.filter.replies',
	reactions: 'activity.filter.reactions',
} as const satisfies Record<ActivityFilter, string>;
const KIND_ICON: Record<ActivityKind, string> = {
	mention: 'at',
	reply: 'reply',
	reaction: 'smile',
	pin: 'pin',
	meeting: 'calendar',
};
const EMPTY_KEY = {
	all: 'activity.empty.all',
	unread: 'activity.empty.unread',
	mentions: 'activity.empty.mentions',
	replies: 'activity.empty.replies',
	reactions: 'activity.empty.reactions',
} as const satisfies Record<ActivityFilter, string>;

/** "14:05" today, "Yesterday" or a short date otherwise. */
export function activityTime(ts: number, now = Date.now(), locale?: string): string {
	const kind = dayKind(ts, now);
	if (kind === 'today') return formatTime(ts, locale);
	const i18n = teamsI18n(locale);
	if (kind === 'yesterday') return i18n.t('time.yesterday');
	return new Date(ts).toLocaleDateString(i18n.intl, { month: 'short', day: 'numeric' });
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
				<h2>${this.t('activity.title')}</h2>
				<button
					type="button"
					class="read-all"
					?disabled=${unread === 0}
					@click=${() => this.fire('office-activity-read-all', {})}
				>
					${this.t('activity.markAllRead')}
				</button>
			</header>
			<div class="filters" role="group" aria-label=${this.t('activity.filterLabel')}>
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
							${this.t(FILTER_KEY[value])}${value === 'unread' && unread ? html` <span class="count">${unread}</span>` : nothing}
						</button>`,
				)}
			</div>
			${
				shown.length
					? html`<ul aria-label=${this.t('activity.feedLabel')}>
							${shown.map((item) => this.row(item))}
						</ul>`
					: html`<p class="empty" role="status">${this.t(EMPTY_KEY[this.filter])}</p>`
			}
		`;
	}

	private row(item: ActivityItem) {
		const line = describeActivityIn(item, this.i18n);
		const toggle = this.t(item.read ? 'activity.markUnread' : 'activity.markRead');
		return html`<li data-unread=${String(!item.read)} data-kind=${item.kind}>
			<button
				type="button"
				class="item"
				aria-current=${this.selectedId === item.id ? 'true' : 'false'}
				aria-label=${`${item.read ? '' : `${this.t('activity.unreadPrefix')} `}${line}. ${item.preview}`}
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
					<time datetime=${new Date(item.ts).toISOString()}
						>${activityTime(item.ts, Date.now(), this.locale)}</time
					>
					${item.read ? nothing : html`<span class="dot" aria-hidden="true"></span>`}
				</span>
			</button>
			<button
				type="button"
				class="toggle"
				title=${toggle}
				aria-label=${toggle}
				@click=${() => this.fire('office-activity-read', { id: item.id, read: !item.read })}
			>
				${icon(item.read ? 'bell' : 'check')}
			</button>
		</li>`;
	}
}

export const defineActivityFeed = definer('office-ui-activity-feed', () => OfficeUiActivityFeed);
