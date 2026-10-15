import { html, nothing } from 'lit';
import { type MeetingView } from 'ooxml-core/teams';
import { formatWhen } from '../base';
import { icon } from '../icons';
import { dialogStyles, TeamsDialog } from './dialog-base';
import css from './meeting-details.css?raw';

const REPEAT_TEXT = { none: '', daily: 'Repeats every day', weekly: 'Repeats every week' } as const;

/**
 * A meeting's details with its actions. Properties: `open`, `meeting`, `occurrenceStart` (the
 * occurrence that was opened), `zone` (the viewer's), `channelName`. Emits `teams-meeting-join`,
 * `teams-meeting-chat`, `teams-meeting-edit`, `teams-meeting-cancel`, `teams-meeting-ics`
 * (all `{ meetingId }`) and `teams-dialog-close`.
 */
export class TeamsMeetingDetails extends TeamsDialog {
	static override styles = dialogStyles(css);
	static override properties = {
		...TeamsDialog.properties,
		meeting: { attribute: false },
		occurrenceStart: { type: Number },
		zone: { type: String },
		channelName: { type: String },
	};
	declare meeting: MeetingView | null;
	declare occurrenceStart: number;
	declare zone: string;
	declare channelName: string;

	constructor() {
		super();
		this.meeting = null;
		this.occurrenceStart = 0;
		this.zone = '';
		this.channelName = '';
	}

	private emit(name: string): void {
		if (this.meeting) this.fire(name, { meetingId: this.meeting.id });
	}

	protected override render() {
		const m = this.meeting;
		if (!m) return this.frame('Meeting', html``);
		const start = this.occurrenceStart || m.start;
		const end = start + (m.end - m.start);
		const hasChat = Boolean(m.chatId ?? m.channelId);
		return this.frame(
			m.title,
			html`<div class="body">
				<p class="when">${icon('calendar')} ${formatWhen(start, end)}</p>
				${m.repeat !== 'none' ? html`<p>${REPEAT_TEXT[m.repeat]}</p>` : nothing}
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
				${m.description ? html`<p class="description">${m.description}</p>` : nothing}
				<footer>
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
							? html`<button type="button" @click=${() => this.emit('teams-meeting-edit')}>
										Edit
									</button>
									<button
										type="button"
										class="danger"
										@click=${() => this.emit('teams-meeting-cancel')}
									>
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
