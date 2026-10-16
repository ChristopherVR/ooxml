import { html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import {
	type PresenceStatus,
	type StatusDuration,
	type TeamsState,
	IDLE_MINUTE_CHOICES,
	MAX_STATUS_MESSAGE_CHARS,
	PRESENCE_STATUSES,
	STATUS_DURATIONS,
	isPresenceStatus,
	isStatusDuration,
} from 'ooxml-core/teams';
import { TeamsLitElement } from '../i18n/lit-element';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';
import { toggleAnchoredPopover } from './anchored-popover';
import css from './profile-menu.css?raw';

const STATUS_KEY = {
	available: 'status.available',
	busy: 'status.busy',
	dnd: 'status.dnd',
	brb: 'status.brb',
	away: 'status.away',
	offline: 'status.offline',
} as const satisfies Record<PresenceStatus, string>;
const CONNECTION_KEY = {
	connecting: 'connection.connecting',
	connected: 'connection.connected',
	disconnected: 'connection.disconnected',
	error: 'connection.error',
} as const satisfies Record<TeamsState['status'], string>;
const DURATION_KEY = {
	never: 'duration.never',
	'30m': 'duration.30m',
	'1h': 'duration.1h',
	'2h': 'duration.2h',
	today: 'duration.today',
	week: 'duration.week',
} as const satisfies Record<StatusDuration, string>;

/** When a status or message ends, as the short sentence under the field ("Until 3:30 PM"). */
export function untilLabel(at: number | null, i18n: TeamsI18n = teamsI18n()): string {
	if (at === null) return '';
	const { t, intl } = i18n;
	const d = new Date(at);
	const today = new Date().toDateString();
	if (d.getHours() === 0 && d.getMinutes() === 0) {
		const lastDay = new Date(at - 1);
		return lastDay.toDateString() === today
			? t('profile.untilEndToday')
			: t('profile.untilEndDay', {
					day: lastDay.toLocaleDateString(intl, { weekday: 'long' }),
				});
	}
	const time = d.toLocaleTimeString(intl, { hour: 'numeric', minute: '2-digit' });
	return d.toDateString() === today
		? t('profile.until', { time })
		: t('profile.untilDayTime', { day: d.toLocaleDateString(intl, { weekday: 'long' }), time });
}

export class TeamsProfileMenu extends TeamsLitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		state: { attribute: false },
		account: { type: String },
		open: { state: true },
		draft: { state: true },
		resetAfter: { state: true },
		clearAfter: { state: true },
	};
	declare state: TeamsState | null;
	/** The signed-in account's user id on a server with accounts; empty otherwise. */
	declare account: string;
	declare open: boolean;
	/** The unsaved text of the status message field. */
	declare draft: string | null;
	/** The chosen reset-after durations, applied with the next status or message. */
	declare resetAfter: StatusDuration;
	declare clearAfter: StatusDuration;
	constructor() {
		super();
		this.state = null;
		this.account = '';
		this.open = false;
		this.draft = null;
		this.resetAfter = 'never';
		this.clearAfter = 'never';
	}
	protected override willUpdate(changed: PropertyValues<this>): void {
		const old = changed.get('state') as TeamsState | null | undefined;
		if (old && old.user.id !== this.state?.user.id)
			this.renderRoot.querySelector<HTMLElement>('[popover]')?.hidePopover();
	}
	private toggle(): void {
		const panel = this.renderRoot.querySelector<HTMLElement>('[popover]')!;
		if (toggleAnchoredPopover(panel, this.renderRoot.querySelector('button')!))
			panel.querySelector('select')?.focus();
	}
	private fire<T>(name: string, detail: T): void {
		this.dispatchEvent(new CustomEvent(name, { detail, bubbles: true, composed: true }));
	}

	private saveMessage(): void {
		this.fire('teams-profile-message', {
			message: this.draft ?? this.state?.presence.message ?? '',
			duration: this.clearAfter,
		});
		this.draft = null;
	}

	private durationOptions(selected: StatusDuration) {
		return STATUS_DURATIONS.map(
			(value) =>
				html`<option value=${value} ?selected=${selected === value}>
					${this.t(DURATION_KEY[value])}
				</option>`,
		);
	}

	private idleLabel(minutes: number): string {
		return minutes === 0 ? this.t('duration.never') : this.tn('profile.idleMinutes', minutes);
	}

	protected override render() {
		const state = this.state;
		if (!state) return nothing;
		const { presence } = state;
		return html`<button
				type="button"
				class="profile"
				aria-label=${this.t('profile.for', { name: state.user.name })}
				aria-expanded=${String(this.open)}
				aria-haspopup="dialog"
				aria-controls="profile"
				@click=${this.toggle}
			>
				<office-ui-avatar
					name=${state.user.name}
					seed=${state.user.id}
					presence=${state.availability}
				></office-ui-avatar>
			</button>
			<div
				id="profile"
				popover="auto"
				role="dialog"
				aria-label=${this.t('profile.yours')}
				@keydown=${(event: KeyboardEvent) => {
					if (event.key !== 'Escape') return;
					event.preventDefault();
					this.renderRoot.querySelector<HTMLElement>('[popover]')?.hidePopover();
					this.renderRoot.querySelector('button')?.focus();
				}}
				@toggle=${() => {
					const panel = this.renderRoot.querySelector<HTMLElement>('[popover]')!;
					this.open = panel.matches(':popover-open');
					if (!this.open && panel.contains(this.shadowRoot?.activeElement ?? null))
						this.renderRoot.querySelector('button')?.focus();
				}}
			>
				<div class="identity">
					<office-ui-avatar
						size="lg"
						name=${state.user.name}
						seed=${state.user.id}
						presence=${state.availability}
					></office-ui-avatar>
					<div class="who">
						<strong>${state.user.name}</strong>
						<span class="status-line"
							>${this.t(STATUS_KEY[state.availability])}${state.presence.idle && state.presence.status === 'available' ? ` ${this.t('profile.idle')}` : ''}</span
						>
						${presence.message ? html`<span class="message">${presence.message}</span>` : nothing}
					</div>
				</div>
				<label
					>${this.t('profile.availability')}<select
						.value=${presence.status}
						@change=${(event: Event) => {
							const availability = (event.target as HTMLSelectElement).value;
							if (!isPresenceStatus(availability)) return;
							this.fire('teams-profile-status', { availability, duration: this.resetAfter });
						}}
					>
						${PRESENCE_STATUSES.map((value) => html`<option value=${value} ?selected=${presence.status === value}>${this.t(STATUS_KEY[value])}</option>`)}
					</select></label
				>
				<label
					>${this.t('profile.resetAfter')}<select
						name="status-duration"
						@change=${(event: Event) => {
							const duration = (event.target as HTMLSelectElement).value;
							if (!isStatusDuration(duration)) return;
							this.resetAfter = duration;
							this.fire('teams-profile-status', { availability: presence.status, duration });
						}}
					>
						${this.durationOptions(this.resetAfter)}
					</select>
					<span class="hint">${untilLabel(presence.statusExpiresAt, this.i18n)}</span></label
				>
				<label
					>${this.t('profile.statusMessage')}<input
						type="text"
						name="message"
						maxlength=${MAX_STATUS_MESSAGE_CHARS}
						placeholder=${this.t('profile.messagePlaceholder')}
						.value=${this.draft ?? presence.message}
						@input=${(event: Event) => (this.draft = (event.target as HTMLInputElement).value)}
						@keydown=${(event: KeyboardEvent) => {
							if (event.key !== 'Enter') return;
							event.preventDefault();
							this.saveMessage();
						}}
					/>
					<span class="hint">${untilLabel(presence.messageExpiresAt, this.i18n)}</span></label
				>
				<label
					>${this.t('profile.clearAfter')}<select
						name="message-duration"
						@change=${(event: Event) => {
							const duration = (event.target as HTMLSelectElement).value;
							if (isStatusDuration(duration)) this.clearAfter = duration;
						}}
					>
						${this.durationOptions(this.clearAfter)}
					</select></label
				>
				<div class="actions">
					<button type="button" @click=${this.saveMessage}>${this.t('profile.saveMessage')}</button>
					<button
						type="button"
						?disabled=${!presence.message && !this.draft}
						@click=${() => {
							this.draft = '';
							this.saveMessage();
						}}
					>
						${this.t('profile.clearMessage')}
					</button>
				</div>
				<label
					>${this.t('profile.awayAfter')}<select
						name="idle"
						@change=${(event: Event) => {
							const minutes = Number((event.target as HTMLSelectElement).value);
							if (IDLE_MINUTE_CHOICES.includes(minutes))
								this.fire('teams-profile-idle', { minutes });
						}}
					>
						${IDLE_MINUTE_CHOICES.map((minutes) => html`<option value=${minutes} ?selected=${presence.idleMinutes === minutes}>${this.idleLabel(minutes)}</option>`)}
					</select></label
				>
				${
					this.account
						? html`<div class="account">
								<span>${this.t('signIn.signedInAs', { id: this.account })}</span>
								<button
									type="button"
									name="sign-out"
									@click=${() => this.fire('teams-sign-out', {})}
								>
									${this.t('signIn.signOut')}
								</button>
							</div>`
						: nothing
				}
				<p class="connection">
					${this.t(state.mode === 'local' ? 'profile.localWorkspace' : 'profile.workspaceServer')}:
					${this.t(CONNECTION_KEY[state.status])}
				</p>
			</div>`;
	}
}

export function defineTeamsProfileMenu(): void {
	if (!customElements.get('teams-profile-menu'))
		customElements.define('teams-profile-menu', TeamsProfileMenu);
}
