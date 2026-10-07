// Personal preferences and connection settings share one navigable settings surface.
// Connection changes are validated by the core before the host applies them.
import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import {
	type TeamsServerConfig,
	type ThreadFollowSettings,
	parseServerConfig,
} from 'ooxml-core/teams';
import css from './teams-settings.css?raw';
import { iceToText, parseIceLines } from './settings-connection.js';
import { densityChoices, type ChatDensity } from './settings-density.js';
import { SETTINGS_CATEGORIES as CATEGORIES, settingsSearch } from './settings-search.js';
export { iceToText, parseIceLines } from './settings-connection.js';

export type TeamsTheme = 'system' | 'light' | 'dark';

/**
 * `<teams-settings .config=${cfg} open>`. Emits `teams-settings-apply` `{ config }` and
 * `teams-settings-close` `{}`.
 */
export class TeamsSettings extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		config: { attribute: false },
		open: { type: Boolean, reflect: true },
		issues: { state: true },
		category: { state: true },
		theme: { attribute: false },
		followSettings: { attribute: false },
		userName: { attribute: false },
		fileOpenPreference: { attribute: false },
		chatDensity: { attribute: false },
		showAppNames: { attribute: false },
		query: { state: true },
	};
	declare config: TeamsServerConfig;
	declare open: boolean;
	declare issues: string[];
	declare category: string;
	declare theme: TeamsTheme;
	declare followSettings: ThreadFollowSettings;
	declare userName: string;
	declare fileOpenPreference: 'teams' | 'browser';
	declare chatDensity: ChatDensity;
	declare showAppNames: boolean;
	declare query: string;

	constructor() {
		super();
		this.config = { mode: 'local', iceServers: [] };
		this.open = false;
		this.issues = [];
		this.category = 'general';
		this.theme = 'system';
		this.followSettings = { started: true, replied: true };
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
		if (this.open && !dialog.open) dialog.showModal();
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
			<dialog aria-label="Settings" @close=${this.close} @cancel=${this.close}>
				<header>
					<h1>Settings</h1>
					<input
						type="search"
						aria-label="Search settings"
						placeholder="Search settings"
						.value=${this.query}
						@input=${(event: Event) => (this.query = (event.target as HTMLInputElement).value)}
					/>
					<button type="button" class="close" aria-label="Close settings" @click=${this.close}>
						×
					</button>
				</header>
				<div class="settings-layout">
					<nav
						aria-label="Settings categories"
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
									${label}
								</button>`,
						)}
					</nav>
					<div class="settings-content">
						${
							this.query.trim()
								? settingsSearch(this.query, async (topic) => {
										this.category = topic.category;
										this.query = '';
										await this.updateComplete;
										if (!this.open || this.query.trim() || this.category !== topic.category) return;
										const panel = this.renderRoot.querySelector<HTMLElement>(
											`#panel-${topic.category}`,
										)!;
										const heading =
											[...panel.querySelectorAll<HTMLElement>('h2, h3, legend')].find(
												(element) => element.textContent?.trim() === topic.title,
											) ?? panel;
										heading.tabIndex = -1;
										heading.focus();
									})
								: nothing
						}
						<section
							id="panel-general"
							role="tabpanel"
							aria-labelledby="settings-general"
							?hidden=${!!this.query.trim() || this.category !== 'general'}
						>
							<h2>General</h2>
							<h3>Profile</h3>
							<p>${this.userName || 'OpenTeams user'}</p>
							<h3>Your workspace</h3>
							<p>${c.mode === 'local' ? 'Local workspace' : 'Connected workspace'}</p>
							<p class="description">
								Preferences are saved for this user and workspace on this device.
							</p>
						</section>
						<section
							id="panel-appearance"
							role="tabpanel"
							aria-labelledby="settings-appearance"
							?hidden=${!!this.query.trim() || this.category !== 'appearance'}
						>
							<h2>Appearance and accessibility</h2>
							<h3>Theme</h3>
							<div class="theme-options">
								${(['system', 'light', 'dark'] as const).map((theme) => html`<button type="button" aria-pressed=${String(this.theme === theme)} @click=${() => this.dispatchEvent(new CustomEvent('teams-settings-theme', { detail: { theme }, bubbles: true, composed: true }))}><span class="theme-swatch" data-theme=${theme}></span>${theme === 'system' ? 'Follow system' : theme === 'light' ? 'Light' : 'Dark'}</button>`)}
							</div>
							${densityChoices(this.chatDensity, (density) => this.dispatchEvent(new CustomEvent('teams-settings-density', { detail: { density }, bubbles: true, composed: true })))}
							<h3>Show app names</h3>
							<label class="preference"
								><span>Display labels beneath app bar icons</span
								><input
									type="checkbox"
									aria-label="Show app names"
									.checked=${this.showAppNames}
									@change=${(event: Event) => this.dispatchEvent(new CustomEvent('teams-settings-app-names', { detail: { show: (event.target as HTMLInputElement).checked }, bubbles: true, composed: true }))}
							/></label>
							<p class="description">
								Changes apply immediately to this OpenTeams workspace. Keyboard focus and system
								high-contrast preferences remain available.
							</p>
						</section>
						<section
							id="panel-notifications"
							role="tabpanel"
							aria-labelledby="settings-notifications"
							?hidden=${!!this.query.trim() || this.category !== 'notifications'}
						>
							<h2>Notifications and activity</h2>
							<h3>Followed threads</h3>
							<p class="description">Choose which conversations appear in Followed threads.</p>
							${(['started', 'replied'] as const).map((key) => html`<label class="preference"><span>${key === 'started' ? 'Threads I start' : 'Threads I reply to'}</span><input type="checkbox" .checked=${this.followSettings[key]} @change=${(event: Event) => this.dispatchEvent(new CustomEvent('teams-settings-follow', { detail: { [key]: (event.target as HTMLInputElement).checked }, bubbles: true, composed: true }))} /></label>`)}
							<p class="description">
								Desktop, email and activity notifications are not available yet.
							</p>
						</section>
						<section
							id="panel-files"
							role="tabpanel"
							aria-labelledby="settings-files"
							?hidden=${!!this.query.trim() || this.category !== 'files'}
						>
							<h2>Files and links</h2>
							<h3>File open preference</h3>
							<label
								>Always open Word, PowerPoint and Excel files in:
								<select
									aria-label="Office file open preference"
									.value=${this.fileOpenPreference}
									@change=${(event: Event) => this.dispatchEvent(new CustomEvent('teams-settings-file-open', { detail: { preference: (event.target as HTMLSelectElement).value }, bubbles: true, composed: true }))}
								>
									<option value="teams">OpenTeams</option>
									<option value="browser">Browser</option>
								</select></label
							>
							<p class="description">
								Browser opens the built-in viewer in a separate tab. Hosts can provide their own
								viewers. Desktop-app defaults and download-location preferences are not available
								yet.
							</p>
						</section>
						<form
							id="panel-connection"
							role="tabpanel"
							aria-labelledby="settings-connection"
							?hidden=${!!this.query.trim() || this.category !== 'connection'}
							@submit=${this.apply}
						>
							<h2>Connection</h2>
							<p class="description">Configure your workspace server and calling service.</p>
							<label>
								<span>Mode</span>
								<select id="mode" .value=${c.mode}>
									<option value="local" ?selected=${c.mode === 'local'}>
										Local: tabs of this browser only (no server)
									</option>
									<option value="server" ?selected=${c.mode === 'server'}>Your server</option>
								</select>
							</label>
							<label>
								<span>Sync URL</span>
								<input
									id="sync"
									.value=${c.syncUrl ?? ''}
									placeholder="ws://localhost:8787/sync"
									spellcheck="false"
									autocomplete="off"
								/>
								<small>y-websocket compatible endpoint; the workspace id is appended.</small>
							</label>
							<label>
								<span>Signaling URL</span>
								<input
									id="signal"
									.value=${c.signalingUrl ?? ''}
									placeholder="ws://localhost:8787/signal"
									spellcheck="false"
									autocomplete="off"
								/>
								<small>WebRTC signaling relay; the call room is appended.</small>
							</label>
							<label>
								<span>Access token</span>
								<input
									id="token"
									type="password"
									.value=${c.token ?? ''}
									placeholder="shared secret"
									autocomplete="off"
								/>
								<small>Sent as ?token= to both sockets. Stored in this browser.</small>
							</label>
							<label>
								<span>ICE servers</span>
								<textarea
									id="ice"
									spellcheck="false"
									.value=${iceToText(c)}
									placeholder="stun:stun.example.org:3478&#10;turn:turn.example.org:3478 user password"
								></textarea>
								<small>One per line: urls [username credential]. TURN needs credentials.</small>
							</label>
							<label>
								<span>ICE transport</span>
								<select id="policy" .value=${c.iceTransportPolicy ?? 'all'}>
									<option value="all">All routes (direct when possible)</option>
									<option value="relay" ?selected=${c.iceTransportPolicy === 'relay'}>
										Relay only (TURN required, hides IP addresses)
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
								<button type="button" @click=${this.close}>Cancel</button>
								<button type="submit" class="primary">Apply and reconnect</button>
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
	if (registry && !registry.get('teams-settings')) registry.define('teams-settings', TeamsSettings);
}
