import { html, nothing, type PropertyValues } from 'lit';
import { type MeetingView, occurrenceAt } from 'ooxml-core/teams';
import { formatWhen } from '../base';
import { icon } from '../icons';
import { dialogStyles, TeamsDialog } from './dialog-base';
import css from './meeting-details.css?raw';

const REPEAT_TEXT = { none: '', daily: 'Repeats every day', weekly: 'Repeats every week' } as const;

/** What an edit or cancel applies to in a repeating meeting. */
export type MeetingScope = 'occurrence' | 'series';

/** The detail of `teams-meeting-edit` and `teams-meeting-cancel`. */
export interface MeetingChange {
	meetingId: string;
	scope: MeetingScope;
	/** The original start of the opened occurrence. */
	recurrenceId: number;
}

/**
 * A meeting's details with its actions. Properties: `open`, `meeting`, `occurrenceStart` (the
 * occurrence that was opened), `zone` (the viewer's), `channelName`. Emits `teams-meeting-join`,
 * `teams-meeting-chat`, `teams-meeting-ics` (all `{ meetingId }`), `teams-meeting-edit` and
 * `teams-meeting-cancel` (`{ meetingId, scope, recurrenceId }`; a repeating meeting first asks
 * whether the change is for this event or the series) and `teams-dialog-close`.
 */
export class TeamsMeetingDetails extends TeamsDialog {
	static override styles = dialogStyles(css);
	static override properties = {
		...TeamsDialog.properties,
		meeting: { attribute: false },
		occurrenceStart: { type: Number },
		zone: { type: String },
		channelName: { type: String },
		pending: { state: true },
		scope: { state: true },
	};
	declare meeting: MeetingView | null;
	declare occurrenceStart: number;
	declare zone: string;
	declare channelName: string;
	declare pending: 'edit' | 'cancel' | null;
	declare scope: MeetingScope;

	constructor() {
		super();
		this.meeting = null;
		this.occurrenceStart = 0;
		this.zone = '';
		this.channelName = '';
		this.pending = null;
		this.scope = 'occurrence';
	}

	protected override willUpdate(changed: PropertyValues<this>): void {
		if (changed.has('open') || changed.has('meeting') || changed.has('occurrenceStart'))
			this.pending = null;
	}

	private emit(name: string): void {
		if (this.meeting) this.fire(name, { meetingId: this.meeting.id });
	}

	private recurrenceId(): number {
		const m = this.meeting;
		return m ? (occurrenceAt(m, this.occurrenceStart || m.start)?.recurrenceId ?? m.start) : 0;
	}

	/** Edit or cancel: a repeating meeting asks which events first. */
	private change(action: 'edit' | 'cancel'): void {
		if (this.meeting?.repeat === 'none') this.finish(action, 'series');
		else {
			this.pending = action;
			this.scope = 'occurrence';
		}
	}

	private finish(action: 'edit' | 'cancel', scope: MeetingScope): void {
		if (!this.meeting) return;
		this.fire(`teams-meeting-${action}`, {
			meetingId: this.meeting.id,
			scope,
			recurrenceId: this.recurrenceId(),
		});
		this.pending = null;
	}

	private endText(m: MeetingView): string {
		const date = (ms: number): string =>
			new Intl.DateTimeFormat([], {
				dateStyle: 'medium',
				...(this.zone ? { timeZone: this.zone } : {}),
			}).format(ms);
		if (m.count !== undefined) return m.count === 1 ? ', 1 time' : `, ${m.count} times`;
		return m.until !== undefined ? `, until ${date(m.until)}` : '';
	}

	private scopeChoice(action: 'edit' | 'cancel') {
		const options: [MeetingScope, string][] = [
			['occurrence', action === 'edit' ? 'This event' : 'This event only'],
			['series', action === 'edit' ? 'All events in the series' : 'All events in the series'],
		];
		return html`<form
			class="scope"
			@submit=${(e: Event) => {
				e.preventDefault();
				this.finish(action, this.scope);
			}}
		>
			<fieldset>
				<legend>
					${action === 'edit' ? 'Edit recurring meeting' : 'Cancel recurring meeting'}
				</legend>
				${options.map(
					([value, label]) =>
						html`<label
							><input
								type="radio"
								name="scope"
								.value=${value}
								.checked=${this.scope === value}
								@change=${() => (this.scope = value)}
							/>
							${label}</label
						>`,
				)}
			</fieldset>
			<footer>
				<button type="button" @click=${() => (this.pending = null)}>Back</button>
				<button type="submit" class=${action === 'cancel' ? 'danger' : 'primary'}>
					${action === 'edit' ? 'Continue' : 'Cancel meeting'}
				</button>
			</footer>
		</form>`;
	}

	protected override render() {
		const m = this.meeting;
		if (!m) return this.frame('Meeting', html``);
		const occurrence = occurrenceAt(m, this.occurrenceStart || m.start);
		const start = occurrence?.start ?? (this.occurrenceStart || m.start);
		const end = occurrence?.end ?? start + (m.end - m.start);
		const description = occurrence?.description ?? m.description;
		const hasChat = Boolean(m.chatId ?? m.channelId);
		return this.frame(
			occurrence?.title ?? m.title,
			html`<div class="body">
				<p class="when">${icon('calendar')} ${formatWhen(start, end)}</p>
				${m.repeat !== 'none' ? html`<p>${REPEAT_TEXT[m.repeat]}${this.endText(m)}</p>` : nothing}
				${
					m.zone !== this.zone
						? html`<p>Scheduled in the organizer's time zone, ${m.zone.replaceAll('_', ' ')}.</p>`
						: nothing
				}
				<p>Organized by ${m.mine ? 'you' : m.organizerName}</p>
				${
					m.channelId
						? html`<p>
								Channel meeting in # ${this.channelName || 'a channel'}. Everyone in the channel can
								join.
							</p>`
						: html`<div class="people">
								<h3>Attendees (${m.attendees.length + 1})</h3>
								<ul>
									<li>${m.organizerName} (organizer)</li>
									${m.attendees.map((a) => html`<li>${a.name}</li>`)}
								</ul>
							</div>`
				}
				${description ? html`<p class="description">${description}</p>` : nothing}
				${this.pending ? this.scopeChoice(this.pending) : nothing}
				<footer ?hidden=${this.pending !== null}>
					<button type="button" class="primary" @click=${() => this.emit('teams-meeting-join')}>
						${m.live ? 'Join now' : 'Join'}
					</button>
					${
						hasChat
							? html`<button type="button" @click=${() => this.emit('teams-meeting-chat')}>
									Meeting chat
								</button>`
							: nothing
					}
					<button type="button" @click=${() => this.emit('teams-meeting-ics')}>Export .ics</button>
					${
						m.mine
							? html`<button type="button" @click=${() => this.change('edit')}>Edit</button>
									<button type="button" class="danger" @click=${() => this.change('cancel')}>
										Cancel meeting
									</button>`
							: nothing
					}
				</footer>
			</div>`,
			'Close meeting details',
		);
	}
}

export function defineTeamsMeetingDetails(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (registry && !registry.get('teams-meeting-details'))
		registry.define('teams-meeting-details', TeamsMeetingDetails);
}
