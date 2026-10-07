import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import type { TeamsState } from 'ooxml-core/teams';
import { toggleAnchoredPopover } from './anchored-popover';
import css from './profile-menu.css?raw';

const AVAILABILITY = ['available', 'busy', 'away'] as const;

export class TeamsProfileMenu extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = { state: { attribute: false }, open: { state: true } };
	declare state: TeamsState | null;
	declare open: boolean;
	constructor() {
		super();
		this.state = null;
		this.open = false;
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
	protected override render() {
		const state = this.state;
		if (!state) return nothing;
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
					></office-ui-avatar
					><strong>${state.user.name}</strong>
				</div>
				<label
					>Availability<select
						.value=${state.availability}
						@change=${(event: Event) => {
							const availability = (event.target as HTMLSelectElement).value;
							if (AVAILABILITY.some((value) => value === availability))
								this.dispatchEvent(
									new CustomEvent('teams-profile-status', {
										detail: { availability },
										bubbles: true,
										composed: true,
									}),
								);
						}}
					>
						${AVAILABILITY.map((value) => html`<option value=${value} ?selected=${state.availability === value}>${value[0]!.toUpperCase() + value.slice(1)}</option>`)}
					</select></label
				>
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
