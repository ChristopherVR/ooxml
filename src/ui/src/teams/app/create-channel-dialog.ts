import { html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import { TeamsLitElement } from '../i18n/lit-element';
import {
	MAX_CHANNEL_NAME_CHARS,
	MAX_CHANNEL_TOPIC_CHARS,
	sanitizeChannelName,
} from 'ooxml-core/teams';
import css from './create-channel-dialog.css?raw';

/** Standard channel creation. The host owns the client and checks unsaved Office content. */
export class TeamsCreateChannelDialog extends TeamsLitElement {
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
			this.issue = this.t('channel.enterName');
			this.renderRoot.querySelector<HTMLInputElement>('input')?.focus();
			return;
		}
		if (!this.create?.(name, this.description)) this.issue = this.t('channel.notCreated');
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
				<h2 id="title">${this.t('channel.create')}</h2>
				<button type="button" aria-label=${this.t('channel.closeCreate')} @click=${this.close}>
					×
				</button>
			</header>
			<form @submit=${this.submit}>
				<label
					>${this.t('channel.name')}<input
						autofocus
						maxlength=${MAX_CHANNEL_NAME_CHARS}
						.value=${this.name}
						@input=${(event: Event) => {
							this.name = (event.target as HTMLInputElement).value;
							this.issue = '';
						}}
						placeholder=${this.t('channel.namePlaceholder')}
						aria-describedby="channel-name-help"
				/></label>
				<small id="channel-name-help"
					>${this.t('channel.maxChars', { count: MAX_CHANNEL_NAME_CHARS })}</small
				>
				<label
					>${this.t('channel.description')}<textarea
						maxlength=${MAX_CHANNEL_TOPIC_CHARS}
						.value=${this.description}
						@input=${(event: Event) => (this.description = (event.target as HTMLTextAreaElement).value)}
						placeholder=${this.t('channel.descriptionPlaceholder')}
					></textarea>
				</label>
				<div>
					<strong>${this.t('channel.standard')}</strong>
					<p>${this.t('channel.standardNote')}</p>
				</div>
				${this.issue ? html`<p class="issue" role="alert">${this.issue}</p>` : nothing}
				<footer>
					<button type="button" @click=${this.close}>${this.t('common.cancel')}</button
					><button type="submit" class="primary">${this.t('common.create')}</button>
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
