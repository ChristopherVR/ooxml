import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import {
	MAX_CHANNEL_NAME_CHARS,
	MAX_CHANNEL_TOPIC_CHARS,
	sanitizeChannelName,
} from 'ooxml-core/teams';
import css from './create-channel-dialog.css?raw';

/** Standard channel creation. The host owns the client and checks unsaved Office content. */
export class TeamsCreateChannelDialog extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		open: { type: Boolean, reflect: true },
		create: { attribute: false },
		name: { state: true },
		description: { state: true },
		issue: { state: true },
	};
	declare open: boolean;
	declare create: ((name: string, description: string) => boolean) | undefined;
	declare name: string;
	declare description: string;
	declare issue: string;
	constructor() {
		super();
		this.open = false;
		this.create = undefined;
		this.name = '';
		this.description = '';
		this.issue = '';
	}
	protected override willUpdate(changed: PropertyValues<this>): void {
		if (changed.has('open') && this.open) {
			this.name = '';
			this.description = '';
			this.issue = '';
		}
	}
	protected override updated(): void {
		const dialog = this.renderRoot.querySelector('dialog')!;
		if (this.open && !dialog.open) dialog.showModal();
		else if (!this.open && dialog.open) dialog.close();
	}
	private close(): void {
		this.dispatchEvent(
			new CustomEvent('teams-create-channel-close', { bubbles: true, composed: true }),
		);
	}
	private submit(event: Event): void {
		event.preventDefault();
		const name = sanitizeChannelName(this.name);
		if (!name) {
			this.issue = 'Enter a channel name.';
			this.renderRoot.querySelector<HTMLInputElement>('input')?.focus();
			return;
		}
		if (!this.create?.(name, this.description))
			this.issue = 'The channel was not created. Your details are kept here.';
	}
	protected override render() {
		return html`<dialog
			aria-labelledby="title"
			@cancel=${(event: Event) => {
				event.preventDefault();
				this.close();
			}}
		>
			<header>
				<h2 id="title">Create a channel</h2>
				<button type="button" aria-label="Close create channel" @click=${this.close}>×</button>
			</header>
			<form @submit=${this.submit}>
				<label
					>Channel name<input
						autofocus
						maxlength=${MAX_CHANNEL_NAME_CHARS}
						.value=${this.name}
						@input=${(event: Event) => {
							this.name = (event.target as HTMLInputElement).value;
							this.issue = '';
						}}
						placeholder="Give your channel a name"
						aria-describedby="channel-name-help"
				/></label>
				<small id="channel-name-help">${MAX_CHANNEL_NAME_CHARS} characters maximum.</small>
				<label
					>Description (optional)<textarea
						maxlength=${MAX_CHANNEL_TOPIC_CHARS}
						.value=${this.description}
						@input=${(event: Event) => (this.description = (event.target as HTMLTextAreaElement).value)}
						placeholder="Help others understand what this channel is for"
					></textarea>
				</label>
				<div>
					<strong>Standard channel</strong>
					<p>
						Available to everyone in this workspace. Private and shared channels are not available
						yet.
					</p>
				</div>
				${this.issue ? html`<p class="issue" role="alert">${this.issue}</p>` : nothing}
				<footer>
					<button type="button" @click=${this.close}>Cancel</button
					><button type="submit" class="primary">Create</button>
				</footer>
			</form>
		</dialog>`;
	}
}
export function defineTeamsCreateChannelDialog(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (registry && !registry.get('teams-create-channel-dialog'))
		registry.define('teams-create-channel-dialog', TeamsCreateChannelDialog);
}
