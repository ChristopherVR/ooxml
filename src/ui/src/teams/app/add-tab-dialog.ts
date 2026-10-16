import { html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import { TeamsLitElement } from '../i18n/lit-element';
import {
	contentUrl,
	detectContentKind,
	filterFiles,
	type ChannelTab,
	type FileEntry,
	type TabContent,
	type TeamsClient,
} from 'ooxml-core/teams';
import css from './add-tab-dialog.css?raw';

const APPS = [
	{ id: 'docx', label: 'app.word', kind: 'docx', letter: 'W' },
	{ id: 'xlsx', label: 'app.excel', kind: 'xlsx', letter: 'X' },
	{ id: 'pptx', label: 'app.powerpoint', kind: 'pptx', letter: 'P' },
	{ id: 'vsdx', label: 'app.visio', kind: 'vsdx', letter: 'V' },
	{ id: 'markdown', label: 'app.markdown', kind: 'markdown', letter: 'M' },
	{ id: 'site', label: 'app.site', kind: 'website', letter: 'H' },
	{ id: 'text', label: 'app.text', kind: 'text', letter: 'T' },
	{ id: 'website', label: 'app.website', kind: null, letter: '↗' },
] as const;
type AppId = (typeof APPS)[number]['id'];
type AddTab = (
	name: string,
	content: TabContent,
	channelId: string,
	client: TeamsClient,
	postToChannel: boolean,
) => ChannelTab | 'canceled' | null;
const fileKey = (file: FileEntry): string => JSON.stringify([file.messageId, file.name, file.url]);

/** Supported apps configure existing channel content. Creation stays with the owning host/core. */
export class TeamsAddTabDialog extends TeamsLitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		open: { type: Boolean },
		client: { attribute: false },
		channelId: { attribute: false },
		channelName: { attribute: false },
		files: { attribute: false },
		add: { attribute: false },
		app: { state: true },
		query: { state: true },
		selectedKey: { state: true },
		name: { state: true },
		url: { state: true },
		error: { state: true },
		postToChannel: { state: true },
	};
	declare open: boolean;
	declare client: TeamsClient | null;
	declare channelId: string;
	declare channelName: string;
	declare files: FileEntry[];
	declare add: AddTab | null;
	declare app: AppId | null;
	declare query: string;
	declare selectedKey: string;
	declare name: string;
	declare url: string;
	declare error: string;
	declare postToChannel: boolean;
	private owner: TeamsClient | null = null;
	private context = '';
	private contextName = '';
	constructor() {
		super();
		this.open = false;
		this.client = null;
		this.channelId = '';
		this.channelName = '';
		this.files = [];
		this.add = null;
		this.app = null;
		this.query = '';
		this.selectedKey = '';
		this.name = '';
		this.url = '';
		this.error = '';
		this.postToChannel = true;
	}
	protected override willUpdate(changed: PropertyValues<this>): void {
		if (changed.has('open') && this.open) {
			this.owner = this.client;
			this.context = this.channelId;
			this.contextName = this.channelName;
			this.app = null;
			this.query = '';
			this.selectedKey = '';
			this.name = '';
			this.url = '';
			this.error = '';
			this.postToChannel = true;
		}
	}
	protected override updated(changed: PropertyValues<this>): void {
		const dialog = this.renderRoot.querySelector('dialog')!;
		if (this.open && !dialog.open) dialog.showModal();
		else if (!this.open && dialog.open) dialog.close();
		if (changed.has('app') && this.open)
			this.renderRoot
				.querySelector<HTMLElement>(this.app ? '[data-configure]' : '[data-first-app]')
				?.focus();
	}
	private close(): void {
		this.dispatchEvent(new CustomEvent('teams-add-tab-close', { bubbles: true, composed: true }));
	}
	private get current(): boolean {
		return (
			!!this.owner &&
			this.owner === this.client &&
			this.context === this.channelId &&
			this.owner.getState().selectedChannelId === this.context &&
			this.owner.getState().channels.some((channel) => channel.id === this.context)
		);
	}
	private eligible(kind: string): FileEntry[] {
		return this.files.filter(
			(file) =>
				detectContentKind(file.name, file.mime) === kind &&
				!!contentUrl(file.url, this.ownerDocument.baseURI),
		);
	}
	private get selected(): FileEntry | undefined {
		const app = APPS.find((entry) => entry.id === this.app);
		return app?.kind
			? this.eligible(app.kind).find((file) => fileKey(file) === this.selectedKey)
			: undefined;
	}
	private availableLabel(kind: string): string {
		return this.tn('addTab.sharedFiles', this.eligible(kind).length);
	}
	private choose(app: AppId): void {
		this.app = app;
		this.query = '';
		this.selectedKey = '';
		this.name = '';
		this.url = '';
		this.error = '';
	}
	private submit(event: SubmitEvent): void {
		event.preventDefault();
		this.error = '';
		if (!this.current || !this.owner || !this.add) {
			this.error = this.t('addTab.channelChanged');
			return;
		}
		if (this.app === 'website' && !contentUrl(this.url, this.ownerDocument.baseURI)) {
			this.error = this.t('addTab.invalidUrl');
			return;
		}
		const content: TabContent | null =
			this.app === 'website'
				? { type: 'website', url: this.url }
				: this.selected
					? { type: 'file', attachment: this.selected }
					: null;
		if (!content) {
			this.error = this.t('addTab.chooseShared');
			return;
		}
		const result = this.add(this.name, content, this.context, this.owner, this.postToChannel);
		if (!result) this.error = this.t('addTab.invalid');
	}
	protected override render() {
		const app = APPS.find((entry) => entry.id === this.app);
		const files = app?.kind ? filterFiles(this.eligible(app.kind), this.query) : [];
		return html`<dialog
			aria-label=${this.t('addTab.title')}
			@cancel=${this.close}
			@close=${this.close}
		>
			<header>
				<h1>${this.t('addTab.title')}</h1>
				<button type="button" aria-label=${this.t('addTab.close')} @click=${this.close}>×</button>
			</header>
			${
				!app
					? html`<section>
							<h2>${this.t('addTab.chooseApp')}</h2>
							<p>${this.t('addTab.pinTo', { channel: this.contextName })}</p>
							<div class="apps">
								${APPS.map(
									(entry, index) =>
										html`<button
											type="button"
											class="app"
											aria-label=${this.t('addTab.choose', { app: this.t(entry.label) })}
											?data-first-app=${index === 0}
											?autofocus=${index === 0}
											@click=${() => this.choose(entry.id)}
										>
											<span class="badge" data-kind=${entry.kind ?? 'other'}>${entry.letter}</span
											><strong>${this.t(entry.label)}</strong>
											<small
												>${entry.kind ? this.availableLabel(entry.kind) : this.t('addTab.addAddress')}</small
											>
										</button>`,
								)}
							</div>
							<p class="hint">${this.t('addTab.fileHint')}</p>
						</section>`
					: html`<form @submit=${this.submit}>
							<div class="app-heading">
								<span class="badge" data-kind=${app.kind ?? 'other'}>${app.letter}</span>
								<h2>${this.t(app.label)}</h2>
								<button
									type="button"
									@click=${() => {
										this.app = null;
										this.error = '';
									}}
								>
									${this.t('common.back')}
								</button>
							</div>
							<label
								>${this.t('addTab.name')}<input
									data-configure
									name="name"
									maxlength="80"
									required
									.value=${this.name}
									@input=${(event: Event) => {
										this.name = (event.target as HTMLInputElement).value;
										this.error = '';
									}}
							/></label>
							${
								!app.kind
									? html`<label
												>${this.t('addTab.websiteUrl')}<input
													type="url"
													required
													maxlength="2048"
													placeholder="https://example.com"
													.value=${this.url}
													@input=${(event: Event) => {
														this.url = (event.target as HTMLInputElement).value;
														this.error = '';
													}}
											/></label>
											<p class="hint">${this.t('addTab.embedHint')}</p>`
									: html`<label
												>${this.t('addTab.searchFiles')}<input
													type="search"
													.value=${this.query}
													@input=${(event: Event) => (this.query = (event.target as HTMLInputElement).value)}
											/></label>
											<fieldset class="files">
												<legend>${this.t('addTab.chooseFile')}</legend>
												${files.map(
													(file) =>
														html`<label class="file"
															><input
																type="radio"
																name="shared-file"
																.value=${fileKey(file)}
																.checked=${this.selectedKey === fileKey(file)}
																@change=${() => {
																	this.selectedKey = fileKey(file);
																	this.name = file.name;
																	this.error = '';
																}}
															/><span
																><strong>${file.name}</strong
																><small
																	>${this.t('addTab.sharedBy', { author: file.author })}</small
																></span
															></label
														>`,
												)}
												${files.length ? nothing : html`<p class="hint">${this.query ? this.t('files.noMatch') : this.t('addTab.noFiles', { app: this.t(app.label) })}</p>`}
											</fieldset>
											${this.selected ? html`<p class="hint">${this.t('addTab.selected', { name: this.selected.name })}</p>` : this.selectedKey ? html`<p role="alert">${this.t('addTab.noLongerShared')}</p>` : nothing}`
							}
							${!this.current ? html`<p role="alert">${this.t('addTab.channelChanged')}</p>` : nothing}
							${this.error ? html`<p role="alert">${this.error}</p>` : nothing}
							<label class="post-choice"
								><input
									type="checkbox"
									.checked=${this.postToChannel}
									@change=${(event: Event) => (this.postToChannel = (event.target as HTMLInputElement).checked)}
								/>${this.t('addTab.postToChannel')}</label
							>
							<footer>
								<button type="button" @click=${this.close}>${this.t('common.cancel')}</button
								><button
									type="submit"
									class="primary"
									?disabled=${!this.current || !this.name.trim() || (app.kind ? !this.selected : !this.url.trim())}
								>
									${this.t('common.save')}
								</button>
							</footer>
						</form>`
			}
		</dialog>`;
	}
}
export function defineTeamsAddTabDialog(): void {
	if (!customElements.get('teams-add-tab-dialog'))
		customElements.define('teams-add-tab-dialog', TeamsAddTabDialog);
}
