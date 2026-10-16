// Personal preferences and connection settings share one navigable settings surface.
// Connection changes are validated by the core before the host applies them.
import { html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import { TeamsLitElement } from '../i18n/lit-element';
import {
	type NotificationSettings,
	type TeamsServerConfig,
	type ThreadFollowSettings,
	parseServerConfig,
} from 'ooxml-core/teams';
import css from './teams-settings.css?raw';
import previewCss from './settings-previews.css?raw';
import { icon } from '../icons';
import { themeChoices } from './settings-theme';
import { defineSwitch, type OfficeUiSwitch } from '../../form/switch';
import { iceToText, parseIceLines } from './settings-connection';
import { densityChoices, type ChatDensity } from './settings-density';
import { desktopPermission, notificationChoices } from './settings-notifications';
import { SETTINGS_CATEGORIES as CATEGORIES, settingsSearch } from './settings-search';
export { iceToText, parseIceLines } from './settings-connection';

export type TeamsTheme = 'system' | 'light' | 'dark';

/**
 * `<teams-settings .config=${cfg} open>`. Emits `teams-settings-apply` `{ config }` and
 * `teams-settings-close` `{}`.
 */
export class TeamsSettings extends TeamsLitElement {
	static override styles = [unsafeCSS(css), unsafeCSS(previewCss)];
	static override properties = {
		config: { attribute: false },
		open: { type: Boolean, reflect: true },
		embedded: { type: Boolean, reflect: true },
		issues: { state: true },
		category: { state: true },
		theme: { attribute: false },
		followSettings: { attribute: false },
		notificationSettings: { attribute: false },
		userName: { attribute: false },
		fileOpenPreference: { attribute: false },
		chatDensity: { attribute: false },
		showAppNames: { attribute: false },
		query: { state: true },
	};
	declare config: TeamsServerConfig;
	declare open: boolean;
	/** Render a modeless settings page inside the host workspace. */
	declare embedded: boolean;
	declare issues: string[];
	declare category: string;
	declare theme: TeamsTheme;
	declare followSettings: ThreadFollowSettings;
	declare notificationSettings: NotificationSettings;
	declare userName: string;
	declare fileOpenPreference: 'teams' | 'browser';
	declare chatDensity: ChatDensity;
	declare showAppNames: boolean;
	declare query: string;

	constructor() {
		super();
		this.config = { mode: 'local', iceServers: [] };
		this.open = false;
		this.embedded = false;
		this.issues = [];
		this.category = 'general';
		this.theme = 'system';
		this.followSettings = { started: true, replied: true };
		this.notificationSettings = { level: 'all', sound: true };
		this.userName = '';
		this.fileOpenPreference = 'teams';
		this.chatDensity = 'comfy';
		this.showAppNames = true;
		this.query = '';
	}
	protected override willUpdate(changed: PropertyValues<this>): void {
		if (changed.has('open') && this.open) this.query = '';
	}

	protected override updated(changed: PropertyValues<this>): void {
		if (changed.has('category')) {
			const content = this.renderRoot.querySelector('.settings-content');
			if (content) content.scrollTop = 0;
		}
		const dialog = this.renderRoot.querySelector('dialog');
		if (!dialog) return;
		if (changed.has('open') && this.open) {
			const c = this.config;
			for (const [id, value] of Object.entries({
				mode: c.mode,
				sync: c.syncUrl ?? '',
				signal: c.signalingUrl ?? '',
				token: c.token ?? '',
				ice: iceToText(c),
				policy: c.iceTransportPolicy ?? 'all',
			}))
				this.field(id).value = value;
		}
		if (this.open && !dialog.open) {
			if (this.embedded) dialog.show();
			else dialog.showModal();
			this.renderRoot.querySelector<HTMLInputElement>('input[type="search"]')?.focus();
		}
		if (!this.open && dialog.open) dialog.close();
	}

	private field(id: string) {
		return this.renderRoot.querySelector<
			HTMLInputElement & HTMLSelectElement & HTMLTextAreaElement
		>(`#${id}`)!;
	}

	private apply(event: Event): void {
		event.preventDefault();
		const { config, issues } = parseServerConfig({
			mode: this.field('mode').value,
			syncUrl: this.field('sync').value,
			signalingUrl: this.field('signal').value,
			token: this.field('token').value,
			iceServers: parseIceLines(this.field('ice').value),
			iceTransportPolicy: this.field('policy').value,
		});
		this.issues = issues;
		// A usable config is applied even with warnings; an unusable one keeps the dialog open.
		if (this.field('mode').value === 'server' && config.mode === 'local') return;
		this.dispatchEvent(
			new CustomEvent('teams-settings-apply', {
				detail: { config },
				bubbles: true,
				composed: true,
			}),
		);
	}

	private close(): void {
		this.issues = [];
		this.dispatchEvent(new CustomEvent('teams-settings-close', { bubbles: true, composed: true }));
	}

	protected override render() {
		const c = this.config;
		return html`
			<dialog
				aria-label=${this.t('settings.title')}
				@close=${() => {
					if (this.open) this.close();
				}}
				@cancel=${(event: Event) => {
					event.preventDefault();
					this.close();
				}}
				@keydown=${(event: KeyboardEvent) => {
					if (this.embedded && event.key === 'Escape' && !event.defaultPrevented) {
						event.preventDefault();
						this.close();
					}
				}}
			>
				<header>
					<h1>${this.t('settings.title')}</h1>
					<input
						type="search"
						aria-label=${this.t('settings.search')}
						placeholder=${this.t('settings.search')}
						.value=${this.query}
						@input=${(event: Event) => (this.query = (event.target as HTMLInputElement).value)}
					/>
					<button
						type="button"
						class="close"
						aria-label=${this.t('settings.close')}
						@click=${this.close}
					>
						×
					</button>
				</header>
				<div class="settings-layout">
					<nav
						aria-label=${this.t('settings.categories')}
						role="tablist"
						aria-orientation="vertical"
						@keydown=${(event: KeyboardEvent) => {
							const index = CATEGORIES.findIndex(([id]) => id === this.category);
							const next =
								event.key === 'ArrowDown'
									? (index + 1) % CATEGORIES.length
									: event.key === 'ArrowUp'
										? (index + CATEGORIES.length - 1) % CATEGORIES.length
										: event.key === 'Home'
											? 0
											: event.key === 'End'
												? CATEGORIES.length - 1
												: -1;
							if (next < 0) return;
							event.preventDefault();
							this.category = CATEGORIES[next]![0];
							this.query = '';
							this.renderRoot
								.querySelector<HTMLButtonElement>(`#settings-${this.category}`)
								?.focus();
						}}
					>
						${CATEGORIES.map(
							([id, label]) =>
								html`<button
									id=${`settings-${id}`}
									role="tab"
									type="button"
									aria-selected=${String(!this.query.trim() && this.category === id)}
									tabindex=${this.category === id ? 0 : -1}
									aria-controls=${`panel-${id}`}
									@click=${() => {
										this.category = id;
										this.query = '';
									}}
								>
									${icon({ general: 'settings', appearance: 'formatPainter', notifications: 'message', files: 'folder', connection: 'screenShare' }[id])}<span
										>${this.t(label)}</span
									>
								</button>`,
						)}
					</nav>
					<div class="settings-content">
						${
							this.query.trim()
								? settingsSearch(
										this.query,
										async (topic) => {
											this.category = topic.category;
											this.query = '';
											await this.updateComplete;
											if (!this.open || this.query.trim() || this.category !== topic.category)
												return;
											const panel = this.renderRoot.querySelector<HTMLElement>(
												`#panel-${topic.category}`,
											)!;
											const heading =
												[...panel.querySelectorAll<HTMLElement>('h2, h3, legend')].find(
													(element) => element.textContent?.trim() === topic.title,
												) ?? panel;
											heading.tabIndex = -1;
											heading.focus();
										},
										this.i18n,
									)
								: nothing
						}
						<section
							id="panel-general"
							role="tabpanel"
							aria-labelledby="settings-general"
							?hidden=${!!this.query.trim() || this.category !== 'general'}
						>
							<h2>${this.t('settings.general')}</h2>
							<h3>${this.t('settings.profile')}</h3>
							<p>${this.userName || this.t('settings.defaultUser')}</p>
							<h3>${this.t('settings.workspace')}</h3>
							<p>
								${this.t(c.mode === 'local' ? 'profile.localWorkspace' : 'settings.connectedWorkspace')}
							</p>
							<p class="description">${this.t('settings.savedNote')}</p>
						</section>
						<section
							id="panel-appearance"
							role="tabpanel"
							aria-labelledby="settings-appearance"
							?hidden=${!!this.query.trim() || this.category !== 'appearance'}
						>
							<h2>${this.t('settings.appearance')}</h2>
							<h3>${this.t('settings.theme')}</h3>
							${themeChoices(this.theme, (theme) => this.dispatchEvent(new CustomEvent('teams-settings-theme', { detail: { theme }, bubbles: true, composed: true })), this.i18n)}
							${densityChoices(this.chatDensity, (density) => this.dispatchEvent(new CustomEvent('teams-settings-density', { detail: { density }, bubbles: true, composed: true })), this.i18n)}
							<h3>${this.t('settings.appNames')}</h3>
							<label class="preference"
								><span>${this.t('settings.appNamesText')}</span
								><office-ui-switch
									aria-label=${this.t('settings.appNames')}
									.checked=${this.showAppNames}
									@change=${(event: Event) => this.dispatchEvent(new CustomEvent('teams-settings-app-names', { detail: { show: (event.target as OfficeUiSwitch).checked }, bubbles: true, composed: true }))}
								></office-ui-switch
							></label>
							<p class="description">${this.t('settings.appearanceNote')}</p>
						</section>
						<section
							id="panel-notifications"
							role="tabpanel"
							aria-labelledby="settings-notifications"
							?hidden=${!!this.query.trim() || this.category !== 'notifications'}
						>
							<h2>${this.t('settings.notifications')}</h2>
							<h3>${this.t('settings.activityNotifications')}</h3>
							${notificationChoices(
								this.notificationSettings,
								desktopPermission(),
								{
									change: (detail) =>
										this.dispatchEvent(
											new CustomEvent('teams-settings-notifications', {
												detail,
												bubbles: true,
												composed: true,
											}),
										),
									requestPermission: () =>
										void globalThis.Notification?.requestPermission().then(() =>
											this.requestUpdate(),
										),
								},
								this.i18n,
							)}
							<h3>${this.t('settings.followed')}</h3>
							<p class="description">${this.t('settings.followedNote')}</p>
							${(['started', 'replied'] as const).map((key) => html`<label class="preference"><span>${this.t(key === 'started' ? 'followed.started' : 'followed.replied')}</span><office-ui-switch aria-label=${this.t(key === 'started' ? 'followed.started' : 'followed.replied')} .checked=${this.followSettings[key]} @change=${(event: Event) => this.dispatchEvent(new CustomEvent('teams-settings-follow', { detail: { [key]: (event.target as OfficeUiSwitch).checked }, bubbles: true, composed: true }))}></office-ui-switch></label>`)}
						</section>
						<section
							id="panel-files"
							role="tabpanel"
							aria-labelledby="settings-files"
							?hidden=${!!this.query.trim() || this.category !== 'files'}
						>
							<h2>${this.t('settings.files')}</h2>
							<h3>${this.t('settings.fileOpen')}</h3>
							<label
								>${this.t('settings.fileOpenLabel')}
								<select
									aria-label=${this.t('settings.fileOpenAria')}
									.value=${this.fileOpenPreference}
									@change=${(event: Event) => this.dispatchEvent(new CustomEvent('teams-settings-file-open', { detail: { preference: (event.target as HTMLSelectElement).value }, bubbles: true, composed: true }))}
								>
									<option value="teams">OpenTeams</option>
									<option value="browser">${this.t('settings.browser')}</option>
								</select></label
							>
							<p class="description">${this.t('settings.fileOpenNote')}</p>
						</section>
						<form
							id="panel-connection"
							role="tabpanel"
							aria-labelledby="settings-connection"
							?hidden=${!!this.query.trim() || this.category !== 'connection'}
							@submit=${this.apply}
						>
							<h2>${this.t('settings.connection')}</h2>
							<p class="description">${this.t('settings.connectionNote')}</p>
							<label>
								<span>${this.t('settings.mode')}</span>
								<select id="mode" .value=${c.mode}>
									<option value="local" ?selected=${c.mode === 'local'}>
										${this.t('settings.modeLocal')}
									</option>
									<option value="server" ?selected=${c.mode === 'server'}>
										${this.t('settings.modeServer')}
									</option>
								</select>
							</label>
							<label>
								<span>${this.t('settings.syncUrl')}</span>
								<input
									id="sync"
									.value=${c.syncUrl ?? ''}
									placeholder="ws://localhost:8787/sync"
									spellcheck="false"
									autocomplete="off"
								/>
								<small>${this.t('settings.syncHint')}</small>
							</label>
							<label>
								<span>${this.t('settings.signalUrl')}</span>
								<input
									id="signal"
									.value=${c.signalingUrl ?? ''}
									placeholder="ws://localhost:8787/signal"
									spellcheck="false"
									autocomplete="off"
								/>
								<small>${this.t('settings.signalHint')}</small>
							</label>
							<label>
								<span>${this.t('settings.token')}</span>
								<input
									id="token"
									type="password"
									.value=${c.token ?? ''}
									placeholder=${this.t('settings.tokenPlaceholder')}
									autocomplete="off"
								/>
								<small>${this.t('settings.tokenHint')}</small>
							</label>
							<label>
								<span>${this.t('settings.iceServers')}</span>
								<textarea
									id="ice"
									spellcheck="false"
									.value=${iceToText(c)}
									placeholder="stun:stun.example.org:3478&#10;turn:turn.example.org:3478 user password"
								></textarea>
								<small>${this.t('settings.iceHint')}</small>
							</label>
							<label>
								<span>${this.t('settings.iceTransport')}</span>
								<select id="policy" .value=${c.iceTransportPolicy ?? 'all'}>
									<option value="all">${this.t('settings.iceAll')}</option>
									<option value="relay" ?selected=${c.iceTransportPolicy === 'relay'}>
										${this.t('settings.iceRelay')}
									</option>
								</select>
							</label>
							${
								this.issues.length
									? html`<ul class="issues">
											${this.issues.map((i) => html`<li>${i}</li>`)}
										</ul>`
									: nothing
							}
							<div class="actions">
								<button type="button" @click=${this.close}>${this.t('common.cancel')}</button>
								<button type="submit" class="primary">${this.t('settings.apply')}</button>
							</div>
						</form>
					</div>
				</div>
			</dialog>
		`;
	}
}

export function defineTeamsSettings(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (!registry) return;
	defineSwitch(registry);
	if (!registry.get('teams-settings')) registry.define('teams-settings', TeamsSettings);
}
