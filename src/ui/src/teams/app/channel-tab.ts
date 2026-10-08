import { LitElement, html, unsafeCSS, type PropertyValues } from 'lit';
import type { ChannelTab, TeamsClient } from 'ooxml-core/teams';
import type { FileEmbeds, SaveFileCopy } from './content-preview';
import type { OpenFileDetail } from './teams-app';
import css from './channel-tab.css?raw';

/** Resolves shared tab references to local, short-lived preview URLs. */
export class TeamsChannelTab extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		tab: { attribute: false },
		client: { attribute: false },
		embeds: { attribute: false },
		canSave: { type: Boolean },
		detail: { state: true },
		error: { state: true },
		hostOpened: { state: true },
	};
	declare tab: ChannelTab | null;
	declare client: TeamsClient | null;
	declare embeds: FileEmbeds;
	declare canSave: boolean;
	declare detail: OpenFileDetail | null;
	declare error: string;
	declare hostOpened: boolean;
	private key = '';
	private owner: TeamsClient | null = null;
	private generation = 0;

	constructor() {
		super();
		this.tab = null;
		this.client = null;
		this.embeds = {};
		this.canSave = false;
		this.detail = null;
		this.error = '';
		this.hostOpened = false;
	}
	protected override updated(changed: PropertyValues<this>): void {
		if (changed.has('tab') || changed.has('client')) void this.resolve();
	}
	override disconnectedCallback(): void {
		this.generation++;
		this.key = '';
		this.owner = null;
		super.disconnectedCallback();
	}
	override connectedCallback(): void {
		super.connectedCallback();
		if (this.hasUpdated) void this.resolve();
	}
	private async resolve(): Promise<void> {
		const { tab, client } = this;
		const key = tab ? `${tab.id}:${JSON.stringify(tab.content)}` : '';
		if (key === this.key && this.owner === client) return;
		this.key = key;
		this.owner = client;
		const generation = ++this.generation;
		this.detail = null;
		this.error = '';
		this.hostOpened = false;
		if (!tab || !client) return;
		const attachment =
			tab.content.type === 'file'
				? tab.content.attachment
				: { name: tab.name, kind: 'other' as const, mime: 'text/html', url: tab.content.url };
		try {
			const url = tab.content.type === 'file' ? await client.fileUrl(attachment) : attachment.url;
			if (generation !== this.generation || !this.isConnected) return;
			if (!url) throw new Error('Could not resolve the tab content');
			this.detail = { attachment, url, channelId: tab.channelId };
			if (tab.content.type === 'file') this.openInHost();
		} catch (error) {
			if (generation === this.generation)
				this.error = error instanceof Error ? error.message : 'Could not open the tab';
		}
	}
	private saveCopy(): SaveFileCopy | undefined {
		const { client, tab } = this;
		if (!this.canSave || !client || !tab) return undefined;
		return async (file, options) => {
			await client.saveFileCopy(tab.channelId, file, options);
		};
	}
	private openInHost(): void {
		this.hostOpened = !this.dispatchEvent(
			new CustomEvent('teams-open-file', {
				detail: this.detail,
				bubbles: true,
				composed: true,
				cancelable: true,
			}),
		);
	}
	protected override render() {
		if (this.hostOpened)
			return html`<p>${this.detail?.attachment.name}</p>
				<button type="button" @click=${() => this.openInHost()}>Open document</button>`;
		return html`${
			this.error
				? html`<p role="alert">${this.error}</p>
						<button
							type="button"
							@click=${() => {
								this.key = '';
								void this.resolve();
							}}
						>
							Retry
						</button>`
				: this.detail
					? html`<teams-content-preview
							.detail=${this.detail}
							.embeds=${this.embeds}
							.saveCopy=${this.saveCopy()}
						></teams-content-preview>`
					: html`<p role="status">Opening tab...</p>`
		}`;
	}
}

export function defineTeamsChannelTab(): void {
	if (globalThis.customElements && !customElements.get('teams-channel-tab'))
		customElements.define('teams-channel-tab', TeamsChannelTab);
}
