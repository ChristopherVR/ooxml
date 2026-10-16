import { html, nothing, type PropertyValues } from 'lit';
import {
	type CalendarDay,
	type CalendarMode,
	type MeetingView,
	layoutDay,
	localZone,
	occurrencesBetween,
	shiftAnchor,
	viewDays,
	zonedParts,
	zonedToUtc,
	startOfDay,
} from 'ooxml-core/teams';
import { definer } from '../registry';
import { TeamsElement, withStyles } from './base';
import { icon } from './icons';
import css from './calendar-view.css?raw';

const MODES = [
	{ id: 'workweek', label: 'calendar.workWeek' },
	{ id: 'week', label: 'calendar.week' },
	{ id: 'day', label: 'calendar.day' },
] as const satisfies readonly { id: CalendarMode; label: string }[];
const SLOT_MINUTES = 30;
const FIRST_VISIBLE_HOUR = 8;

/**
 * A Teams-style calendar: a time grid for the work week, the full week or one day in the viewer's
 * zone, with Today, previous and next, and the view switch. Overlapping meetings share a column.
 * Properties: `meetings` (core `MeetingView[]`), `zone` (default: this device's), `now` (the
 * clock, for tests), `mode`. Events: `office-calendar-open` `{ meetingId, start }`,
 * `office-calendar-new` `{ start }` (a click on an empty slot, or the New meeting button), and
 * `office-calendar-mode` `{ mode }`.
 */
export class OfficeUiCalendarView extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		meetings: { attribute: false },
		zone: { type: String },
		now: { type: Number },
		mode: { type: String, reflect: true },
		anchor: { state: true },
		tick: { state: true },
	};
	declare meetings: MeetingView[];
	declare zone: string;
	declare now: number;
	declare mode: CalendarMode;
	declare anchor: number;
	declare tick: number;
	private timer: ReturnType<typeof setInterval> | undefined;
	private scrolled = false;
	private resize: ResizeObserver | undefined;

	constructor() {
		super();
		this.meetings = [];
		this.zone = localZone();
		this.now = 0;
		this.mode = 'workweek';
		this.anchor = 0;
		this.tick = 0;
	}

	private clock(): number {
		return this.now || Date.now();
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.setAttribute('role', 'region');
		if (!this.anchor) this.anchor = this.clock();
		// A phone starts on the day view, like the Teams mobile calendar.
		if (globalThis.matchMedia?.('(max-width: 40em)').matches && this.mode === 'workweek')
			this.mode = 'day';
		this.timer = setInterval(() => (this.tick = this.clock()), 60_000);
		// A phone shows the list first, so the grid may get its size long after it first rendered.
		if (typeof ResizeObserver !== 'undefined') {
			this.resize = new ResizeObserver(() => this.scrollToWorkday());
			this.resize.observe(this);
		}
	}

	protected override hostLabel(): string {
		return this.t('calendar.title');
	}

	override disconnectedCallback(): void {
		clearInterval(this.timer);
		this.resize?.disconnect();
		super.disconnectedCallback();
	}

	protected override updated(): void {
		this.scrollToWorkday();
	}

	/** Start the day at 8:00, once the grid has a size. */
	private scrollToWorkday(): void {
		const scroller = this.renderRoot.querySelector<HTMLElement>('.scroller');
		if (this.scrolled || !scroller || scroller.clientHeight === 0) return;
		scroller.scrollTop = (scroller.scrollHeight / 24) * FIRST_VISIBLE_HOUR;
		this.scrolled = true;
	}

	protected override willUpdate(changed: PropertyValues<this>): void {
		if (changed.has('now') && this.now && changed.get('now') === 0) this.anchor = this.now;
	}

	private setMode(mode: CalendarMode): void {
		this.mode = mode;
		this.fire('office-calendar-mode', { mode });
	}

	private range(days: CalendarDay[]): string {
		const first = days[0]!;
		const last = days[days.length - 1]!;
		const intl = this.i18n.intl;
		const fmt = (opts: Intl.DateTimeFormatOptions, ms: number): string =>
			new Intl.DateTimeFormat(intl, { ...opts, timeZone: this.zone }).format(ms);
		if (days.length === 1)
			return fmt({ weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }, first.start);
		return new Intl.DateTimeFormat(intl, {
			month: 'long',
			day: 'numeric',
			year: 'numeric',
			timeZone: this.zone,
		}).formatRange(first.start, last.start);
	}

	private newAt(day: CalendarDay, event: MouseEvent): void {
		const column = event.currentTarget as HTMLElement;
		const box = column.getBoundingClientRect();
		const fraction = box.height ? (event.clientY - box.top) / box.height : 0;
		const minute = Math.floor((Math.max(0, fraction) * 1440) / SLOT_MINUTES) * SLOT_MINUTES;
		const start = zonedToUtc(
			{ year: day.year, month: day.month, day: day.day, hour: 0, minute },
			this.zone,
		);
		this.fire('office-calendar-new', { start });
	}

	protected override render() {
		const now = this.clock();
		const days = viewDays(this.anchor || now, this.mode, this.zone);
		const todayStart = startOfDay(now, this.zone);
		const nowParts = zonedParts(now, this.zone);
		return html`
			<div class="bar">
				<h2 class="range" aria-live="polite">${this.range(days)}</h2>
				<button type="button" class="today" @click=${() => (this.anchor = this.clock())}>
					${this.t('time.today')}
				</button>
				<button
					type="button"
					class="step"
					aria-label=${this.t(this.mode === 'day' ? 'calendar.previousDay' : 'calendar.previousWeek')}
					@click=${() => (this.anchor = shiftAnchor(this.anchor || now, this.mode, -1, this.zone))}
				>
					${icon('back')}
				</button>
				<button
					type="button"
					class="step next"
					aria-label=${this.t(this.mode === 'day' ? 'calendar.nextDay' : 'calendar.nextWeek')}
					@click=${() => (this.anchor = shiftAnchor(this.anchor || now, this.mode, 1, this.zone))}
				>
					${icon('back')}
				</button>
				<div class="modes" role="group" aria-label=${this.t('calendar.viewLabel')}>
					${MODES.map(
						(m) =>
							html`<button
								type="button"
								aria-pressed=${String(this.mode === m.id)}
								@click=${() => this.setMode(m.id)}
							>
								${this.t(m.label)}
							</button>`,
					)}
				</div>
				<button
					type="button"
					class="new"
					@click=${() => this.fire('office-calendar-new', { start: 0 })}
				>
					${icon('plus')} ${this.t('calendar.newMeeting')}
				</button>
			</div>
			<div class="scroller">
				<div class="grid" style=${`--cols:${days.length}`}>
					<div class="corner"></div>
					${days.map((d) => this.dayHead(d, d.start === todayStart))}
					<div class="hours" aria-hidden="true">
						${Array.from({ length: 24 }, (_, h) => html`<span>${this.hourLabel(h)}</span>`)}
					</div>
					${days.map((d) => this.column(d, d.start === todayStart, nowParts.hour * 60 + nowParts.minute))}
				</div>
			</div>
		`;
	}

	private hourLabel(hour: number): string {
		if (hour === 0) return '';
		return new Date(Date.UTC(2000, 0, 1, hour)).toLocaleTimeString(this.i18n.intl, {
			hour: 'numeric',
			timeZone: 'UTC',
		});
	}

	private dayHead(day: CalendarDay, today: boolean) {
		const name = new Intl.DateTimeFormat(this.i18n.intl, {
			weekday: 'short',
			timeZone: this.zone,
		}).format(day.start);
		return html`<div class="head" data-today=${String(today)}>
			<span class="dow">${name}</span>
			<span class="dom">${day.day}</span>
		</div>`;
	}

	private column(day: CalendarDay, today: boolean, nowMinute: number) {
		const occurrences = this.meetings.flatMap((m) =>
			occurrencesBetween(m, day.start, day.end).map((o) => ({ o, m })),
		);
		const placed = layoutDay(
			occurrences.map((x) => x.o),
			day,
			this.zone,
		);
		const byKey = new Map(occurrences.map((x) => [x.o.key, x.m]));
		const label = new Intl.DateTimeFormat(this.i18n.intl, {
			weekday: 'long',
			month: 'long',
			day: 'numeric',
			timeZone: this.zone,
		}).format(day.start);
		return html`<div
			class="column"
			role="group"
			aria-label=${label}
			data-today=${String(today)}
			@click=${(e: MouseEvent) => {
				if (e.target === e.currentTarget) this.newAt(day, e);
			}}
		>
			${placed.map((p) => {
				const meeting = byKey.get(p.occurrence.key)!;
				const title = p.occurrence.title ?? meeting.title;
				const time = `${this.clockText(p.occurrence.start)} - ${this.clockText(p.occurrence.end)}`;
				return html`<button
					type="button"
					class="event"
					data-live=${String(meeting.live)}
					data-mine=${String(meeting.mine)}
					data-short=${String(p.endMinute - p.startMinute <= 30)}
					aria-label=${`${title}, ${time}`}
					style=${`--top:${(p.startMinute / 1440) * 100};--height:${((p.endMinute - p.startMinute) / 1440) * 100};--col:${p.column};--span:${p.columns}`}
					@click=${(e: Event) => {
						e.stopPropagation();
						this.fire('office-calendar-open', {
							meetingId: meeting.id,
							start: p.occurrence.start,
						});
					}}
				>
					<span class="event-title">${title}</span>
					<span class="event-time">${time}</span>
				</button>`;
			})}
			${
				today
					? html`<span
							class="now"
							style=${`--top:${(nowMinute / 1440) * 100}`}
							aria-hidden="true"
						></span>`
					: nothing
			}
		</div>`;
	}

	private clockText(ms: number): string {
		return new Date(ms).toLocaleTimeString(this.i18n.intl, {
			hour: 'numeric',
			minute: '2-digit',
			timeZone: this.zone,
		});
	}
}

export const defineCalendarView = definer('office-ui-calendar-view', () => OfficeUiCalendarView);
