import { html } from 'lit';
import {
	DAY_MS,
	type MeetingView,
	type Occurrence,
	dayKey,
	localZone,
	occurrencesBetween,
	startOfDay,
} from 'ooxml-core/teams';
import { definer } from '../registry';
import { TeamsElement, withStyles } from './base';
import { icon } from './icons';
import css from './calendar-agenda.css?raw';

const HORIZON_DAYS = 14;
const JOIN_WINDOW_MS = 15 * 60_000;

/**
 * The Calendar app's side list: meetings in the next two weeks grouped by day, each with its time
 * and a Join button once it is near or live. Properties: `meetings`, `zone`, `now` (the clock, for
 * tests). Emits `office-agenda-open` `{ meetingId, start }`, `office-agenda-join` `{ meetingId }`
 * and `office-agenda-new` `{}`.
 */
export class OfficeUiCalendarAgenda extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		meetings: { attribute: false },
		zone: { type: String },
		now: { type: Number },
	};
	declare meetings: MeetingView[];
	declare zone: string;
	declare now: number;

	constructor() {
		super();
		this.meetings = [];
		this.zone = localZone();
		this.now = 0;
	}

	protected override render() {
		const now = this.now || Date.now();
		const from = startOfDay(now, this.zone);
		const rows = this.meetings
			.flatMap((m) =>
				occurrencesBetween(m, from, from + HORIZON_DAYS * DAY_MS).map((o) => ({ o, m })),
			)
			.filter(({ o }) => o.end > now)
			.sort((a, b) => a.o.start - b.o.start);
		const groups = new Map<string, { label: string; rows: typeof rows }>();
		for (const row of rows) {
			const key = dayKey(row.o.start, this.zone);
			let group = groups.get(key);
			if (!group) {
				const label = new Intl.DateTimeFormat([], {
					weekday: 'long',
					month: 'short',
					day: 'numeric',
					timeZone: this.zone,
				}).format(row.o.start);
				groups.set(
					key,
					(group = { label: key === dayKey(now, this.zone) ? `Today, ${label}` : label, rows: [] }),
				);
			}
			group.rows.push(row);
		}
		return html`
			<header>
				<h2>Calendar</h2>
				<button
					class="new"
					type="button"
					title="New meeting"
					aria-label="New meeting"
					@click=${() => this.fire('office-agenda-new', {})}
				>
					${icon('plus')}
				</button>
			</header>
			${
				groups.size === 0
					? html`<p class="empty">No upcoming meetings. Schedule one with the plus button.</p>`
					: [...groups.values()].map(
							(g) =>
								html`<section>
									<h3>${g.label}</h3>
									<ul>
										${g.rows.map(({ o, m }) => this.row(o, m, now))}
									</ul>
								</section>`,
						)
			}
		`;
	}

	private row(o: Occurrence, m: MeetingView, now: number) {
		const time = `${this.clock(o.start)} - ${this.clock(o.end)}`;
		const joinable = m.live || (o.start - now <= JOIN_WINDOW_MS && o.end > now);
		return html`<li>
			<button
				type="button"
				class="row"
				data-live=${String(m.live)}
				aria-label=${`${m.title}, ${time}`}
				@click=${() => this.fire('office-agenda-open', { meetingId: m.id, start: o.start })}
			>
				<span class="time">${time}</span>
				<span class="title">${m.title}</span>
				<span class="who">${m.mine ? 'You organized' : m.organizerName}</span>
			</button>
			${
				joinable
					? html`<button
							type="button"
							class="join"
							aria-label=${`Join ${m.title}`}
							@click=${() => this.fire('office-agenda-join', { meetingId: m.id })}
						>
							${m.live ? 'Join now' : 'Join'}
						</button>`
					: ''
			}
		</li>`;
	}

	private clock(ms: number): string {
		return new Date(ms).toLocaleTimeString([], {
			hour: 'numeric',
			minute: '2-digit',
			timeZone: this.zone,
		});
	}
}

export const defineCalendarAgenda = definer(
	'office-ui-calendar-agenda',
	() => OfficeUiCalendarAgenda,
);
