import { html, nothing, type PropertyValues } from 'lit';
import {
	type ContactView,
	type MeetingFormInput,
	type MeetingView,
	MAX_MEETING_DESCRIPTION,
	MAX_MEETING_TITLE,
	MINUTE_MS,
	MAX_RECURRENCE_COUNT,
	localZone,
	occurrenceAt,
} from 'ooxml-core/teams';
import { dialogStyles, TeamsDialog } from './dialog-base';
import css from './meeting-dialog.css?raw';
import {
	NO_REPEAT,
	type RepeatSelection,
	parseRepeatValue,
	repeatOptions,
	repeatValue,
} from './repeat-text';

const SLOT = 30 * MINUTE_MS;

const repeatOption = (o: { value: string; label: string }, current: string) =>
	html`<option value=${o.value} ?selected=${o.value === current}>${o.label}</option>`;

/** The repeat fields of a stored meeting: a weekday rule carries its ordinal. */
function repeatFields(
	r: RepeatSelection,
): Pick<MeetingFormInput, 'repeat' | 'monthlyBy' | 'ordinal'> {
	if (r.repeat !== 'monthly') return { repeat: r.repeat };
	return r.monthlyBy === 'weekday'
		? { repeat: 'monthly', monthlyBy: 'weekday', ordinal: r.ordinal }
		: { repeat: 'monthly', monthlyBy: 'date' };
}

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
		repeatChoice: { state: true },
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
	declare repeatChoice: RepeatSelection;
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
		this.repeatChoice = NO_REPEAT;
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
		this.repeatChoice = m
			? { repeat: m.repeat, monthlyBy: m.monthlyBy ?? 'date', ordinal: m.ordinal ?? 1 }
			: NO_REPEAT;
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

	/** The zone the series is defined in: the organizer's for an existing meeting, else this device's. */
	private zone(): string {
		return this.meeting?.zone ?? localZone();
	}

	/**
	 * The repeat to store. A weekday ordinal that no longer fits the start date (the start moved)
	 * becomes the first one that does.
	 */
	private repeatToStore(start: number): RepeatSelection {
		const choice = this.repeatChoice;
		if (choice.repeat !== 'monthly' || choice.monthlyBy !== 'weekday') return choice;
		const fits = repeatOptions(start, this.zone(), this.i18n).map((o) => o.value);
		return fits.includes(repeatValue(choice))
			? choice
			: parseRepeatValue(fits.find((v) => v.startsWith('monthly-') && v !== 'monthly-date') ?? '');
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
		if (this.repeatChoice.repeat === 'none' || this.endMode === 'never')
			return { count: null, until: null };
		return this.endMode === 'count'
			? { count: Number(this.countText), until: null }
			: { count: null, until: this.untilTime() };
	}

	private submit(event: Event): void {
		event.preventDefault();
		const start = new Date(this.startText).getTime();
		const end = new Date(this.endText).getTime();
		if (!this.meetingTitle.trim()) this.issue = this.t('meeting.addTitle');
		else if (!Number.isFinite(start) || !Number.isFinite(end))
			this.issue = this.t('meeting.pickStartEnd');
		else if (end <= start) this.issue = this.t('meeting.endAfterStart');
		else if (end - start > 24 * 60 * MINUTE_MS) this.issue = this.t('meeting.maxDuration');
		else if (this.occurrence) {
			const one = { title: this.meetingTitle, description: this.description, start, end };
			if (!this.save?.(one, this.meeting?.id, this.occurrence))
				this.issue = this.t('meeting.notSaved');
		} else if (
			this.repeatChoice.repeat !== 'none' &&
			this.endMode === 'count' &&
			!this.validCount()
		)
			this.issue = this.t('meeting.countRange', { max: MAX_RECURRENCE_COUNT });
		else if (
			this.repeatChoice.repeat !== 'none' &&
			this.endMode === 'date' &&
			!this.validUntil(start)
		)
			this.issue = this.t('meeting.untilBeforeStart');
		else if (
			!this.save?.(
				{
					title: this.meetingTitle,
					description: this.description,
					start,
					end,
					...repeatFields(this.repeatToStore(start)),
					...this.recurrenceEnd(),
					attendeeIds: this.channelId ? [] : this.attendees,
					...(this.channelId ? { channelId: this.channelId } : {}),
				},
				this.meeting?.id,
			)
		)
			this.issue = this.t('meeting.notSaved');
	}

	protected override render() {
		const editing = this.meeting !== null;
		const single = this.occurrence !== 0;
		const current = repeatValue(this.repeatToStore(new Date(this.startText).getTime()));
		return this.frame(
			this.t(single ? 'meeting.editEvent' : editing ? 'meeting.edit' : 'calendar.newMeeting'),
			html`<form @submit=${this.submit}>
				<label
					>${this.t('meeting.title')}<input
						maxlength=${MAX_MEETING_TITLE}
						.value=${this.meetingTitle}
						placeholder=${this.t('meeting.titlePlaceholder')}
						@input=${(e: Event) => {
							this.meetingTitle = (e.target as HTMLInputElement).value;
							this.issue = '';
						}}
				/></label>
				<div class="times">
					<label
						>${this.t('meeting.start')}<input
							type="datetime-local"
							.value=${this.startText}
							@change=${(e: Event) => this.moveStart((e.target as HTMLInputElement).value)}
					/></label>
					<label
						>${this.t('meeting.end')}<input
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
						? html`<small>${this.t('meeting.onlyThisEvent')}</small>`
						: html`
								<label
									>${this.t('meeting.repeat')}<select
										@change=${(e: Event) => (this.repeatChoice = parseRepeatValue((e.target as HTMLSelectElement).value))}
									>
										${repeatOptions(
											new Date(this.startText).getTime() || Date.now(),
											this.zone(),
											this.i18n,
										).map((o) => repeatOption(o, current))}
									</select></label
								>
								${
									this.repeatChoice.repeat === 'none'
										? nothing
										: html`<label
													>${this.t('meeting.ends')}<select
														.value=${this.endMode}
														@change=${(e: Event) => (this.endMode = (e.target as HTMLSelectElement).value as typeof this.endMode)}
													>
														<option value="never">${this.t('meeting.endsNever')}</option>
														<option value="date">${this.t('meeting.endsDate')}</option>
														<option value="count">${this.t('meeting.endsCount')}</option>
													</select></label
												>
												${
													this.endMode === 'date'
														? html`<label
																>${this.t('meeting.endDate')}<input
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
																>${this.t('meeting.numberOfEvents')}<input
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
									>${this.t('meeting.channelOptional')}<select
										?disabled=${editing}
										.value=${this.channelId}
										@change=${(e: Event) => (this.channelId = (e.target as HTMLSelectElement).value)}
									>
										<option value="">${this.t('meeting.noChannel')}</option>
										${this.channels.map((c) => html`<option value=${c.id}># ${c.name}</option>`)}
									</select></label
								>
								${
									this.channelId
										? html`<small>${this.t('meeting.channelNote')}</small>`
										: html`<div class="field">
												${this.t('meeting.attendees')}
												<office-ui-people-picker
													label=${this.t('meeting.attendees')}
													.contacts=${this.contacts}
													.selected=${this.attendees}
													@office-people-change=${(e: CustomEvent<{ selected: string[] }>) => (this.attendees = e.detail.selected)}
												></office-ui-people-picker>
											</div>`
								}
							`
				}
				<label
					>${this.t('meeting.detailsOptional')}<textarea
						maxlength=${MAX_MEETING_DESCRIPTION}
						.value=${this.description}
						placeholder=${this.t('meeting.detailsPlaceholder')}
						@input=${(e: Event) => (this.description = (e.target as HTMLTextAreaElement).value)}
					></textarea>
				</label>
				<small>${this.t('meeting.zoneNote')}</small>
				${this.issue ? html`<p class="issue" role="alert">${this.issue}</p>` : nothing}
				<footer>
					<button type="button" @click=${this.requestClose}>${this.t('common.cancel')}</button>
					<button type="submit" class="primary">
						${this.t(editing ? 'common.save' : 'meeting.send')}
					</button>
				</footer>
			</form>`,
			this.t(
				single
					? 'dialog.closeEditEvent'
					: editing
						? 'dialog.closeEditMeeting'
						: 'dialog.closeNewMeeting',
			),
		);
	}
}

export function defineTeamsMeetingDialog(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (registry && !registry.get('teams-meeting-dialog'))
		registry.define('teams-meeting-dialog', TeamsMeetingDialog);
}
