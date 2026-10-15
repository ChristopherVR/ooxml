import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import {
	type StatusDuration,
	type TeamsState,
	IDLE_MINUTE_CHOICES,
	MAX_STATUS_MESSAGE_CHARS,
	PRESENCE_LABELS,
	PRESENCE_STATUSES,
	STATUS_DURATIONS,
	STATUS_DURATION_LABELS,
	isPresenceStatus,
	isStatusDuration,
} from 'ooxml-core/teams';
import { toggleAnchoredPopover } from './anchored-popover';
import css from './profile-menu.css?raw';
import { SIGN_IN_STRINGS } from './sign-in-strings';

const idleLabel = (minutes: number): string => (minutes === 0 ? 'Never' : `${minutes} minutes`);
/** When a status or message ends, as the short sentence under the field ("Until 3:30 PM"). */
const untilLabel = (at: number | null): string => {
	if (at === null) return '';
	const d = new Date(at);
	const today = new Date().toDateString();
	if (d.getHours() === 0 && d.getMinutes() === 0) {
		const lastDay = new Date(at - 1);
		return lastDay.toDateString() === today
			? 'Until end of today'
			: `Until end of ${lastDay.toLocaleDateString([], { weekday: 'long' })}`;
	}
	const time = d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
	return d.toDateString() === today
		? `Until ${time}`
		: `Until ${d.toLocaleDateString([], { weekday: 'long' })} ${time}`;
};

export class TeamsProfileMenu extends LitElement {
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
					${STATUS_DURATION_LABELS[value]}
				</option>`,
		);
	}

	protected override render() {
		const state = this.state;
		if (!state) return nothing;
		const { presence } = state;
		return html`<button
				type="button"
				class="profile"
				aria-label=${`Profile for ${state.user.name}`}
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
				aria-label="Your profile"
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
							>${PRESENCE_LABELS[state.availability]}${state.presence.idle && state.presence.status === 'available' ? ' (idle)' : ''}</span
						>
						${presence.message ? html`<span class="message">${presence.message}</span>` : nothing}
					</div>
				</div>
				<label
					>Availability<select
						.value=${presence.status}
						@change=${(event: Event) => {
							const availability = (event.target as HTMLSelectElement).value;
							if (!isPresenceStatus(availability)) return;
							this.fire('teams-profile-status', { availability, duration: this.resetAfter });
						}}
					>
						${PRESENCE_STATUSES.map((value) => html`<option value=${value} ?selected=${presence.status === value}>${PRESENCE_LABELS[value]}</option>`)}
					</select></label
				>
				<label
					>Reset status after<select
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
					<span class="hint">${untilLabel(presence.statusExpiresAt)}</span></label
				>
				<label
					>Status message<input
						type="text"
						name="message"
						maxlength=${MAX_STATUS_MESSAGE_CHARS}
						placeholder="What's your status message?"
						.value=${this.draft ?? presence.message}
						@input=${(event: Event) => (this.draft = (event.target as HTMLInputElement).value)}
						@keydown=${(event: KeyboardEvent) => {
							if (event.key !== 'Enter') return;
							event.preventDefault();
							this.saveMessage();
						}}
					/>
					<span class="hint">${untilLabel(presence.messageExpiresAt)}</span></label
				>
				<label
					>Clear status message after<select
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
					<button type="button" @click=${this.saveMessage}>Save message</button>
					<button
						type="button"
						?disabled=${!presence.message && !this.draft}
						@click=${() => {
							this.draft = '';
							this.saveMessage();
						}}
					>
						Clear message
					</button>
				</div>
				<label
					>Show me as away after<select
						name="idle"
						@change=${(event: Event) => {
							const minutes = Number((event.target as HTMLSelectElement).value);
							if (IDLE_MINUTE_CHOICES.includes(minutes))
								this.fire('teams-profile-idle', { minutes });
						}}
					>
						${IDLE_MINUTE_CHOICES.map((minutes) => html`<option value=${minutes} ?selected=${presence.idleMinutes === minutes}>${idleLabel(minutes)}</option>`)}
					</select></label
				>
				${
					this.account
						? html`<div class="account">
								<span>${SIGN_IN_STRINGS.signedInAs(this.account)}</span>
								<button
									type="button"
									name="sign-out"
									@click=${() => this.fire('teams-sign-out', {})}
								>
									${SIGN_IN_STRINGS.signOut}
								</button>
							</div>`
						: nothing
				}
				<p class="connection">
					${state.mode === 'local' ? 'Local workspace' : 'Workspace server'}: ${state.status}
				</p>
			</div>`;
	}
}

export function defineTeamsProfileMenu(): void {
	if (!customElements.get('teams-profile-menu'))
		customElements.define('teams-profile-menu', TeamsProfileMenu);
}
