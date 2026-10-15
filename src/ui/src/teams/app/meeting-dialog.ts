import { html, nothing, type PropertyValues } from 'lit';
import {
	type ContactView,
	type MeetingFormInput,
	type MeetingRepeat,
	type MeetingView,
	MAX_MEETING_DESCRIPTION,
	MAX_MEETING_TITLE,
	MINUTE_MS,
	MAX_RECURRENCE_COUNT,
	occurrenceAt,
} from 'ooxml-core/teams';
import { dialogStyles, TeamsDialog } from './dialog-base';
import css from './meeting-dialog.css?raw';

const SLOT = 30 * MINUTE_MS;
const NOT_SAVED = 'The meeting was not saved. Your details are kept here.';

/** `2026-10-15T14:30`, the value of a datetime-local input, in this device's zone. */
export function toLocalInput(ms: number): string {
	const d = new Date(ms);
	const p = (n: number): string => String(n).padStart(2, '0');
	return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
}

/**
 * Schedule or edit a meeting. Properties: `open`, `meeting` (null to schedule), `start` (the
 * prefilled start, 0 for the next half hour; the opened occurrence when editing one), `occurrence`
 * (the original start of the one occurrence of a repeating meeting to edit: only its time, title
 * and details), `contacts`, `channels` (for a channel meeting), `save(form, id?, occurrence?)`
 * (true when stored). Emits `teams-dialog-close`.
 */
export class TeamsMeetingDialog extends TeamsDialog {
	static override styles = dialogStyles(css);
	static override properties = {
		...TeamsDialog.properties,
		meeting: { attribute: false },
		start: { type: Number },
		occurrence: { type: Number },
		contacts: { attribute: false },
		channels: { attribute: false },
		save: { attribute: false },
		meetingTitle: { state: true },
		description: { state: true },
		startText: { state: true },
		endText: { state: true },
		repeat: { state: true },
		endMode: { state: true },
		untilText: { state: true },
		countText: { state: true },
		channelId: { state: true },
		attendees: { state: true },
		issue: { state: true },
	};
	declare meeting: MeetingView | null;
	declare start: number;
	declare occurrence: number;
	declare contacts: ContactView[];
	declare channels: { id: string; name: string }[];
	declare save: ((form: MeetingFormInput, id?: string, occurrence?: number) => boolean) | undefined;
	declare meetingTitle: string;
	declare description: string;
	declare startText: string;
	declare endText: string;
	declare repeat: MeetingRepeat;
	declare endMode: 'never' | 'date' | 'count';
	declare untilText: string;
	declare countText: string;
	declare channelId: string;
	declare attendees: string[];
	declare issue: string;

	constructor() {
		super();
		this.meeting = null;
		this.start = 0;
		this.occurrence = 0;
		this.contacts = [];
		this.channels = [];
		this.save = undefined;
		this.meetingTitle = '';
		this.description = '';
		this.startText = '';
		this.endText = '';
		this.repeat = 'none';
		this.endMode = 'never';
		this.untilText = '';
		this.countText = '10';
		this.channelId = '';
		this.attendees = [];
		this.issue = '';
	}

	protected override willUpdate(changed: PropertyValues<this>): void {
		if (!changed.has('open') || !this.open) return;
		const m = this.meeting;
		const one = m && this.occurrence ? occurrenceAt(m, this.start || m.start) : null;
		const begin = one?.start ?? m?.start ?? (this.start || Math.ceil(Date.now() / SLOT) * SLOT);
		this.meetingTitle = one?.title ?? m?.title ?? '';
		this.description = one?.description ?? m?.description ?? '';
		this.startText = toLocalInput(begin);
		this.endText = toLocalInput(one?.end ?? m?.end ?? begin + SLOT);
		this.repeat = m?.repeat ?? 'none';
		this.endMode = m?.count !== undefined ? 'count' : m?.until !== undefined ? 'date' : 'never';
		this.countText = String(m?.count ?? 10);
		this.untilText = toLocalInput(m?.until ?? begin + 28 * 86_400_000).slice(0, 10);
		this.channelId = m?.channelId ?? '';
		this.attendees = m?.attendees.map((a) => a.id) ?? [];
		this.issue = '';
	}

	private moveStart(value: string): void {
		const before = new Date(this.startText).getTime();
		const after = new Date(value).getTime();
		const end = new Date(this.endText).getTime();
		this.startText = value;
		// Keep the length of the meeting when its start moves, as calendars do.
		if (Number.isFinite(before) && Number.isFinite(after) && Number.isFinite(end))
			this.endText = toLocalInput(end + (after - before));
		this.issue = '';
	}

	private validCount(): boolean {
		const n = Number(this.countText);
		return Number.isInteger(n) && n >= 1 && n <= MAX_RECURRENCE_COUNT;
	}

	private untilTime(): number {
		return new Date(`${this.untilText}T23:59:59.999`).getTime();
	}

	private validUntil(start: number): boolean {
		const until = this.untilTime();
		return Number.isFinite(until) && until >= start;
	}

	/** The end of the repeat to store: a count, the end of a day, or none (null clears an old end). */
	private recurrenceEnd(): { count: number | null; until: number | null } {
		if (this.repeat === 'none' || this.endMode === 'never') return { count: null, until: null };
		return this.endMode === 'count'
			? { count: Number(this.countText), until: null }
			: { count: null, until: this.untilTime() };
	}

	private submit(event: Event): void {
		event.preventDefault();
		const start = new Date(this.startText).getTime();
		const end = new Date(this.endText).getTime();
		if (!this.meetingTitle.trim()) this.issue = 'Add a title.';
		else if (!Number.isFinite(start) || !Number.isFinite(end)) this.issue = 'Pick a start and end.';
		else if (end <= start) this.issue = 'The meeting must end after it starts.';
		else if (end - start > 24 * 60 * MINUTE_MS) this.issue = 'A meeting can last at most 24 hours.';
		else if (this.occurrence) {
			const one = { title: this.meetingTitle, description: this.description, start, end };
			if (!this.save?.(one, this.meeting?.id, this.occurrence)) this.issue = NOT_SAVED;
		} else if (this.repeat !== 'none' && this.endMode === 'count' && !this.validCount())
			this.issue = `Enter a number of events from 1 to ${MAX_RECURRENCE_COUNT}.`;
		else if (this.repeat !== 'none' && this.endMode === 'date' && !this.validUntil(start))
			this.issue = 'The series must end on or after its first event.';
		else if (
			!this.save?.(
				{
					title: this.meetingTitle,
					description: this.description,
					start,
					end,
					repeat: this.repeat,
					...this.recurrenceEnd(),
					attendeeIds: this.channelId ? [] : this.attendees,
					...(this.channelId ? { channelId: this.channelId } : {}),
				},
				this.meeting?.id,
			)
		)
			this.issue = NOT_SAVED;
	}

	protected override render() {
		const editing = this.meeting !== null;
		const single = this.occurrence !== 0;
		return this.frame(
			single ? 'Edit this event' : editing ? 'Edit meeting' : 'New meeting',
			html`<form @submit=${this.submit}>
				<label
					>Title<input
						maxlength=${MAX_MEETING_TITLE}
						.value=${this.meetingTitle}
						placeholder="Add a title"
						@input=${(e: Event) => {
							this.meetingTitle = (e.target as HTMLInputElement).value;
							this.issue = '';
						}}
				/></label>
				<div class="times">
					<label
						>Start<input
							type="datetime-local"
							.value=${this.startText}
							@change=${(e: Event) => this.moveStart((e.target as HTMLInputElement).value)}
					/></label>
					<label
						>End<input
							type="datetime-local"
							.value=${this.endText}
							@change=${(e: Event) => {
								this.endText = (e.target as HTMLInputElement).value;
								this.issue = '';
							}}
					/></label>
				</div>
				${
					single
						? html`<small>Only this event changes. The rest of the series stays as it is.</small>`
						: html`
								<label
									>Repeat<select
										.value=${this.repeat}
										@change=${(e: Event) => (this.repeat = (e.target as HTMLSelectElement).value as MeetingRepeat)}
									>
										<option value="none">Does not repeat</option>
										<option value="daily">Every day</option>
										<option value="weekly">Every week</option>
									</select></label
								>
								${
									this.repeat === 'none'
										? nothing
										: html`<label
													>Ends<select
														.value=${this.endMode}
														@change=${(e: Event) => (this.endMode = (e.target as HTMLSelectElement).value as typeof this.endMode)}
													>
														<option value="never">Never</option>
														<option value="date">On a date</option>
														<option value="count">After a number of events</option>
													</select></label
												>
												${
													this.endMode === 'date'
														? html`<label
																>End date<input
																	type="date"
																	.value=${this.untilText}
																	@change=${(e: Event) => {
																		this.untilText = (e.target as HTMLInputElement).value;
																		this.issue = '';
																	}}
															/></label>`
														: nothing
												}
												${
													this.endMode === 'count'
														? html`<label
																>Number of events<input
																	type="number"
																	min="1"
																	max=${MAX_RECURRENCE_COUNT}
																	.value=${this.countText}
																	@input=${(e: Event) => {
																		this.countText = (e.target as HTMLInputElement).value;
																		this.issue = '';
																	}}
															/></label>`
														: nothing
												}`
								}
								<label
									>Channel (optional)<select
										?disabled=${editing}
										.value=${this.channelId}
										@change=${(e: Event) => (this.channelId = (e.target as HTMLSelectElement).value)}
									>
										<option value="">No channel: invite people</option>
										${this.channels.map((c) => html`<option value=${c.id}># ${c.name}</option>`)}
									</select></label
								>
								${
									this.channelId
										? html`<small
												>A channel meeting posts a meeting card in the channel and shows in
												everyone's calendar.</small
											>`
										: html`<div class="field">
												Attendees
												<office-ui-people-picker
													label="Attendees"
													.contacts=${this.contacts}
													.selected=${this.attendees}
													@office-people-change=${(e: CustomEvent<{ selected: string[] }>) => (this.attendees = e.detail.selected)}
												></office-ui-people-picker>
											</div>`
								}
							`
				}
				<label
					>Details (optional)<textarea
						maxlength=${MAX_MEETING_DESCRIPTION}
						.value=${this.description}
						placeholder="Add an agenda or notes"
						@input=${(e: Event) => (this.description = (e.target as HTMLTextAreaElement).value)}
					></textarea>
				</label>
				<small
					>Times are shown in your time zone. The organizer's zone keeps repeats on schedule.</small
				>
				${this.issue ? html`<p class="issue" role="alert">${this.issue}</p>` : nothing}
				<footer>
					<button type="button" @click=${this.requestClose}>Cancel</button>
					<button type="submit" class="primary">${editing ? 'Save' : 'Send'}</button>
				</footer>
			</form>`,
		);
	}
}

export function defineTeamsMeetingDialog(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (registry && !registry.get('teams-meeting-dialog'))
		registry.define('teams-meeting-dialog', TeamsMeetingDialog);
}
