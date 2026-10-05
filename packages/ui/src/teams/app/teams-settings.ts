// The "Server" settings dialog: where chat syncs, where calls are signalled and which STUN/TURN
// servers carry media. Input goes through core's `parseServerConfig`, which reports every problem
// instead of throwing; nothing is applied until the form is valid enough to use.
import { LitElement, html, nothing, unsafeCSS } from 'lit';
import { type TeamsServerConfig, parseServerConfig } from 'ooxml-core/teams';
import css from './teams-settings.css?raw';

/** One ICE server per line: `urls[,urls] [username credential]`. */
export function parseIceLines(text: string): unknown[] {
	return text
		.split('\n')
		.map((line) => line.trim())
		.filter(Boolean)
		.map((line) => {
			const [urls = '', username, credential] = line.split(/\s+/u);
			return {
				urls: urls.split(',').filter(Boolean),
				...(username ? { username } : {}),
				...(credential ? { credential } : {}),
			};
		});
}

export const iceToText = (config: TeamsServerConfig): string =>
	config.iceServers
		.map((s) => [[s.urls].flat().join(','), s.username, s.credential].filter(Boolean).join(' '))
		.join('\n');

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
	};
	declare config: TeamsServerConfig;
	declare open: boolean;
	declare issues: string[];

	constructor() {
		super();
		this.config = { mode: 'local', iceServers: [] };
		this.open = false;
		this.issues = [];
	}

	protected override updated(): void {
		const dialog = this.renderRoot.querySelector('dialog');
		if (!dialog) return;
		if (this.open && !dialog.open) dialog.showModal();
		if (!this.open && dialog.open) dialog.close();
	}

	private field(id: string) {
		return this.renderRoot.querySelector<HTMLInputElement & HTMLSelectElement & HTMLTextAreaElement>(`#${id}`)!;
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
		this.dispatchEvent(new CustomEvent('teams-settings-apply', { detail: { config }, bubbles: true, composed: true }));
	}

	private close(): void {
		this.dispatchEvent(new CustomEvent('teams-settings-close', { bubbles: true, composed: true }));
	}

	protected override render() {
		const c = this.config;
		return html`
			<dialog @close=${this.close} @cancel=${this.close}>
				<form @submit=${this.apply}>
					<h2>Server settings</h2>
					<label>
						<span>Mode</span>
						<select id="mode" .value=${c.mode}>
							<option value="local" ?selected=${c.mode === 'local'}>Local: tabs of this browser only (no server)</option>
							<option value="server" ?selected=${c.mode === 'server'}>Your server</option>
						</select>
					</label>
					<label>
						<span>Sync URL</span>
						<input id="sync" .value=${c.syncUrl ?? ''} placeholder="ws://localhost:8787/sync" spellcheck="false" autocomplete="off" />
						<small>y-websocket compatible endpoint; the workspace id is appended.</small>
					</label>
					<label>
						<span>Signaling URL</span>
						<input id="signal" .value=${c.signalingUrl ?? ''} placeholder="ws://localhost:8787/signal" spellcheck="false" autocomplete="off" />
						<small>WebRTC signaling relay; the call room is appended.</small>
					</label>
					<label>
						<span>Access token</span>
						<input id="token" type="password" .value=${c.token ?? ''} placeholder="shared secret" autocomplete="off" />
						<small>Sent as ?token= to both sockets. Stored in this browser.</small>
					</label>
					<label>
						<span>ICE servers</span>
						<textarea id="ice" spellcheck="false" .value=${iceToText(c)} placeholder="stun:stun.example.org:3478&#10;turn:turn.example.org:3478 user password"></textarea>
						<small>One per line: urls [username credential]. TURN needs credentials.</small>
					</label>
					<label>
						<span>ICE transport</span>
						<select id="policy" .value=${c.iceTransportPolicy ?? 'all'}>
							<option value="all">All routes (direct when possible)</option>
							<option value="relay" ?selected=${c.iceTransportPolicy === 'relay'}>Relay only (TURN required, hides IP addresses)</option>
						</select>
					</label>
					${this.issues.length
						? html`<ul class="issues">${this.issues.map((i) => html`<li>${i}</li>`)}</ul>`
						: nothing}
					<div class="actions">
						<button type="button" @click=${this.close}>Cancel</button>
						<button type="submit" class="primary">Apply and reconnect</button>
					</div>
				</form>
			</dialog>
		`;
	}
}

export function defineTeamsSettings(registry: CustomElementRegistry | undefined = globalThis.customElements): void {
	if (registry && !registry.get('teams-settings')) registry.define('teams-settings', TeamsSettings);
}
