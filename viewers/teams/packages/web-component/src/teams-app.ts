// <teams-app>: one reactive element, one template, one CSS file. All collaboration and call logic
// lives in `createTeamsClient` (ooxml-core/teams); the visual pieces are ooxml-ui elements. This
// file only decides what is on screen (`render`) and forwards user intent to client actions.
import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import {
	type Attachment,
	type FileUploader,
	type OfficeKind,
	type TeamsClient,
	type TeamsServerConfig,
	type TeamsState,
	parseServerConfig,
} from 'ooxml-core/teams';
import { registerOfficeUi } from 'ooxml-ui';
import { TeamsController } from './controller.js';
import { defineTeamsSettings } from './teams-settings.js';
import { loadConfig, loadIdentity, safeStorage, saveConfig, saveIdentity, type Identity } from './storage.js';
import css from './teams-app.css?raw';

export type { FileUploader } from 'ooxml-core/teams';
export interface OpenFileDetail {
	attachment: { name: string; kind: OfficeKind; size?: number | undefined; url?: string | undefined };
	/** The URL to open, with the access token appended when the file lives on your server. */
	url: string | undefined;
}
export type FileOpeners = Partial<Record<OfficeKind, (detail: OpenFileDetail) => void>>;

type RailView = 'teams' | 'calls' | 'files';
type Panel = '' | 'chat' | 'people';
const RAIL = [
	{ id: 'teams', label: 'Teams', icon: 'users' },
	{ id: 'calls', label: 'Calls', icon: 'phone' },
	{ id: 'files', label: 'Files', icon: 'folder' },
] as const;
const AVAILABILITY = ['available', 'busy', 'away'] as const;

export class TeamsApp extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		workspaceId: { type: String, attribute: 'workspace-id' },
		userName: { type: String, attribute: 'user-name' },
		userId: { type: String, attribute: 'user-id' },
		serverConfig: { type: String, attribute: 'server-config' },
		config: { attribute: false },
		uploadFile: { attribute: false },
		openers: { attribute: false },
		rail: { state: true },
		tab: { state: true },
		panel: { state: true },
		meeting: { state: true },
		settingsOpen: { state: true },
		toast: { state: true },
		identity: { state: true },
	};
	declare workspaceId: string;
	declare userName: string;
	declare userId: string;
	declare serverConfig: string;
	/** Server settings. Omit to use the settings dialog (remembered in this browser) or local mode. */
	declare config: TeamsServerConfig | null;
	/** Where attachments go. Default: your server's `/files` endpoint. */
	declare uploadFile: FileUploader | undefined;
	/** Per-product openers (for example the Word or Excel viewer). Default: open the URL. */
	declare openers: FileOpeners;
	declare rail: RailView;
	declare tab: 'posts' | 'files';
	declare panel: Panel;
	declare meeting: boolean;
	declare settingsOpen: boolean;
	declare toast: string;
	declare identity: Identity | null;

	private readonly teams = new TeamsController(this, (text) => this.notify(text));
	private toastTimer: ReturnType<typeof setTimeout> | undefined;

	constructor() {
		super();
		this.workspaceId = 'demo';
		this.userName = '';
		this.userId = '';
		this.serverConfig = '';
		this.config = null;
		this.uploadFile = undefined;
		this.openers = {};
		this.rail = 'teams';
		this.tab = 'posts';
		this.panel = '';
		this.meeting = false;
		this.settingsOpen = false;
		this.toast = '';
		this.identity = null;
	}

	/** The core client behind this element, for hosts that want the raw actions. */
	get client(): TeamsClient | null {
		return this.teams.client;
	}

	override connectedCallback(): void {
		registerOfficeUi();
		defineTeamsSettings();
		super.connectedCallback();
	}

	// ---- lifecycle --------------------------------------------------------------------------
	private resolveConfig(): TeamsServerConfig {
		if (this.config) return this.config;
		if (this.serverConfig) {
			try {
				return parseServerConfig(JSON.parse(this.serverConfig)).config;
			} catch {
				// An invalid attribute falls through to the remembered settings.
			}
		}
		return loadConfig() ?? parseServerConfig({}).config;
	}

	protected override willUpdate(changed: PropertyValues<this>): void {
		const restart = ['workspaceId', 'userName', 'userId', 'serverConfig', 'config', 'identity'].some((k) => changed.has(k as never));
		if (!restart && this.teams.client) return;
		const named = this.userName
			? { id: this.userId || loadIdentity()?.id || crypto.randomUUID(), name: this.userName }
			: this.identity ?? loadIdentity();
		if (this.userName && !this.userId) saveIdentity(named!);
		if (!named) return this.teams.stop();
		const workspaceId = /^[\w-]{1,100}$/u.test(this.workspaceId) ? this.workspaceId : 'demo';
		this.teams.start({
			workspaceId,
			user: named,
			config: this.resolveConfig(),
			storage: safeStorage,
			...(this.uploadFile ? { uploadFile: this.uploadFile } : {}),
		});
		this.dispatchEvent(new CustomEvent('teams-ready', { detail: { user: named }, bubbles: true, composed: true }));
	}

	private notify(text: string): void {
		this.toast = text;
		clearTimeout(this.toastTimer);
		this.toastTimer = setTimeout(() => (this.toast = ''), 4000);
	}

	// ---- intent -> client actions -----------------------------------------------------------
	private async openFile(a: { name: string; kind: OfficeKind; size?: number | undefined; url?: string | undefined }): Promise<void> {
		const url = await this.teams.client?.fileUrl(a);
		const detail: OpenFileDetail = { attachment: a, url };
		if (!this.dispatchEvent(new CustomEvent('teams-open-file', { detail, bubbles: true, composed: true, cancelable: true }))) return;
		const opener = this.openers[a.kind];
		if (opener) return opener(detail);
		if (url) globalThis.open(url, '_blank', 'noopener');
		else this.notify('This file was shared by name only');
	}

	private createChannel(): void {
		const name = globalThis.prompt?.('Channel name')?.trim();
		if (name) {
			this.teams.client?.createChannel(name);
			this.rail = 'teams';
			this.meeting = false;
		}
	}

	private onCallAction(action: string): void {
		const c = this.teams.client;
		if (!c) return;
		if (action === 'chat' || action === 'people') this.panel = this.panel === action ? '' : action;
		else if (action === 'leave') {
			this.meeting = false;
			this.panel = '';
			c.leaveCall();
		} else if (action === 'mic') c.toggleMic();
		else if (action === 'camera') void c.toggleCamera();
		else if (action === 'screen') void c.toggleScreenShare();
		else if (action === 'hand') c.toggleHand();
	}

	private startMeeting(channelId?: string): void {
		void this.teams.client?.openCall(channelId);
		this.meeting = true;
		this.rail = 'teams';
	}

	private selectChannel(id: string): void {
		this.teams.client?.select(id);
		this.teams.client?.search('');
		this.meeting = false;
		this.tab = 'posts';
		this.rail = 'teams';
	}

	private askSettings(): void {
		this.settingsOpen = true;
	}

	// ---- template ---------------------------------------------------------------------------
	protected override render() {
		const s = this.teams.state;
		if (!s) return this.welcome();
		return html`
			<div class="shell">
				${this.topbar(s)}
				<div class="body">
					<office-ui-app-rail
						.items=${RAIL.map((i) => ({ ...i, ...(i.id === 'calls' && s.channels.some((c) => c.live) ? { badge: s.channels.filter((c) => c.live).length } : {}) }))}
						selected=${this.rail}
						@office-rail-select=${(e: CustomEvent<{ id: RailView }>) => ((this.rail = e.detail.id), (this.meeting = false))}
					></office-ui-app-rail>
					<aside class="side">${this.sidebar(s)}</aside>
					<main class="main">${this.main(s)}</main>
				</div>
				${this.toast ? html`<div class="toast" role="status">${this.toast}</div>` : nothing}
				<teams-settings
					?open=${this.settingsOpen}
					.config=${this.resolveConfig()}
					@teams-settings-close=${() => (this.settingsOpen = false)}
					@teams-settings-apply=${(e: CustomEvent<{ config: TeamsServerConfig }>) => {
						saveConfig(e.detail.config);
						this.config = e.detail.config;
						this.settingsOpen = false;
						this.dispatchEvent(new CustomEvent('teams-config-change', { detail: e.detail, bubbles: true, composed: true }));
					}}
				></teams-settings>
			</div>
		`;
	}

	private welcome() {
		return html`
			<div class="welcome">
				<form
					@submit=${(e: SubmitEvent) => {
						e.preventDefault();
						const name = new FormData(e.target as HTMLFormElement).get('name')?.toString().trim();
						if (!name) return;
						const identity = { id: crypto.randomUUID(), name };
						saveIdentity(identity);
						this.identity = identity;
					}}
				>
					<span class="logo">T</span>
					<h1>Welcome to Teams</h1>
					<p>Pick the name your colleagues will see.</p>
					<input name="name" placeholder="Your name" aria-label="Your name" maxlength="40" autocomplete="name" required />
					<button type="submit">Continue</button>
				</form>
			</div>
		`;
	}

	private topbar(s: TeamsState) {
		return html`
			<header class="topbar">
				<div class="brand"><span class="logo">T</span><span>Teams</span></div>
				<div class="search">
					<input
						type="search"
						placeholder="Search messages"
						aria-label="Search messages"
						.value=${s.searchQuery}
						@input=${(e: Event) => this.teams.client?.search((e.target as HTMLInputElement).value)}
					/>
					${s.searchQuery
						? html`<ul class="results" role="listbox" aria-label="Search results">
								${s.searchResults.length === 0
									? html`<li class="none">No results</li>`
									: s.searchResults.map(
											(h) => html`<li>
												<button type="button" @click=${() => this.selectChannel(h.channelId)}>
													<strong>${h.message.authorName}</strong>
													<span class="in">in # ${h.channelName}</span>
													<span class="snippet">${h.message.text.slice(0, 100)}</span>
												</button>
											</li>`,
										)}
							</ul>`
						: nothing}
				</div>
				<div class="me">
					<span class="conn" data-status=${s.status} title=${s.mode === 'server' ? `Server: ${s.status}` : 'Local: tabs of this browser'}></span>
					<button type="button" class="link" @click=${this.askSettings}>Server</button>
					<select
						aria-label="Availability"
						.value=${s.availability}
						@change=${(e: Event) => this.teams.client?.setAvailability((e.target as HTMLSelectElement).value as 'available')}
					>
						${AVAILABILITY.map((a) => html`<option value=${a} ?selected=${a === s.availability}>${a[0]!.toUpperCase() + a.slice(1)}</option>`)}
					</select>
					<office-ui-avatar name=${s.user.name} seed=${s.user.id} presence=${s.availability}></office-ui-avatar>
				</div>
			</header>
		`;
	}

	private sidebar(s: TeamsState) {
		if (this.rail === 'calls') {
			const live = s.channels.filter((c) => c.live);
			return html`
				<h2>Calls</h2>
				<ul class="plain">
					${s.channels.map(
						(c) => html`<li>
							<span># ${c.name}</span>
							<button type="button" class=${c.live ? 'join live' : 'join'} @click=${() => this.startMeeting(c.id)}>
								${c.live ? 'Join' : 'Meet now'}
							</button>
						</li>`,
					)}
				</ul>
				${live.length === 0 ? html`<p class="hint">Nobody is in a call right now.</p>` : nothing}
			`;
		}
		return html`
			<office-ui-channel-list
				heading="Channels"
				selected-id=${s.selectedChannelId}
				.channels=${s.channels}
				@office-channel-select=${(e: CustomEvent<{ id: string }>) => this.selectChannel(e.detail.id)}
				@office-channel-create=${this.createChannel}
			></office-ui-channel-list>
		`;
	}

	private main(s: TeamsState) {
		if (this.rail === 'files') return this.allFiles(s);
		if (s.call && this.meeting) return this.meetingView(s);
		return html`
			${s.call && !this.meeting
				? html`<div class="callbar">
						You are in a call in # ${s.call.channelName}
						<button type="button" @click=${() => ((this.meeting = true), this.teams.client?.select(s.call!.channelId))}>Return to call</button>
					</div>`
				: nothing}
			${this.conversation(s)}
		`;
	}

	private conversation(s: TeamsState, compact = false) {
		const channel = s.channel;
		const c = this.teams.client;
		const compose = s.editing ?? s.replyingTo;
		return html`
			${compact
				? nothing
				: html`<header class="channel-head">
						<h1><span class="hash">#</span> ${channel?.name ?? 'No channel yet'}</h1>
						<nav class="tabs" role="tablist">
							${(['posts', 'files'] as const).map(
								(t) => html`<button
									type="button"
									role="tab"
									aria-selected=${String(this.tab === t)}
									@click=${() => (this.tab = t)}
								>
									${t === 'posts' ? 'Posts' : 'Files'}
								</button>`,
							)}
						</nav>
						<span class="topic">${channel?.topic ?? ''}</span>
						<div class="people" aria-label="People online">
							${s.people.slice(0, 5).map((p) => html`<office-ui-avatar size="sm" name=${p.name} seed=${p.id} color=${p.color} presence=${p.availability}></office-ui-avatar>`)}
							${s.people.length > 5 ? html`<span class="more">+${s.people.length - 5}</span>` : nothing}
						</div>
						<button type="button" class="meet" ?disabled=${!channel} @click=${() => this.startMeeting()}>
							${s.channels.find((x) => x.id === s.selectedChannelId)?.live ? 'Join call' : 'Meet now'}
						</button>
					</header>`}
			${this.tab === 'files' && !compact
				? this.fileList(s.files)
				: html`
						<office-ui-chat-list
							.messages=${s.messages}
							self-id=${s.user.id}
							@office-chat-react=${(e: CustomEvent<{ messageId: string; emoji: string }>) => c?.toggleReaction(e.detail.messageId, e.detail.emoji)}
							@office-chat-reply=${(e: CustomEvent<{ messageId: string }>) => c?.startReply(e.detail.messageId)}
							@office-chat-edit=${(e: CustomEvent<{ messageId: string }>) => c?.startEdit(e.detail.messageId)}
							@office-chat-delete=${(e: CustomEvent<{ messageId: string }>) => globalThis.confirm?.('Delete this message?') !== false && c?.deleteMessage(e.detail.messageId)}
							@office-chat-open-file=${(e: CustomEvent<{ attachment: OpenFileDetail['attachment'] & { kind: OfficeKind } }>) => this.openFile(e.detail.attachment)}
						></office-ui-chat-list>
						<office-ui-chat-composer
							.typing=${s.typing}
							.replyingTo=${s.replyingTo?.authorName ?? null}
							.editing=${s.editing !== null}
							.value=${s.editing?.text ?? ''}
							?disabled=${!channel}
							placeholder=${channel ? `Message # ${channel.name}` : 'Create a channel first'}
							@office-chat-typing=${() => c?.notifyTyping()}
							@office-chat-cancel=${() => c?.cancelCompose()}
							@office-chat-send=${(e: CustomEvent<{ text: string; files: File[] }>) => void c?.send(e.detail)}
							data-compose=${compose ? 'on' : 'off'}
						></office-ui-chat-composer>
					`}
		`;
	}

	private fileList(files: TeamsState['files']) {
		return html`
			<ul class="files">
				${files.length === 0
					? html`<li class="none">No files shared yet.</li>`
					: files.map(
							(f) => html`<li>
								<span class="badge" data-kind=${f.kind}>${f.kind === 'other' ? 'F' : f.kind[0]!.toUpperCase()}</span>
								<span class="meta">
									<strong>${f.name}</strong>
									<small>${f.author} · ${new Date(f.ts).toLocaleString()}${f.channelName ? ` · # ${f.channelName}` : ''}</small>
								</span>
								<button type="button" @click=${() => this.openFile(f)}>Open</button>
							</li>`,
						)}
			</ul>
		`;
	}

	private allFiles(s: TeamsState) {
		return html`<header class="channel-head"><h1>Files</h1></header>${this.fileList(s.allFiles)}`;
	}

	private meetingView(s: TeamsState) {
		const call = s.call!;
		if (call.phase === 'prejoin')
			return html`
				<office-ui-prejoin
					heading=${`# ${call.channelName}`}
					name=${s.user.name}
					.preview=${call.preview as unknown as MediaStream | null}
					?audio=${call.prejoin.audio}
					?video=${call.prejoin.video}
					@office-prejoin-change=${(e: CustomEvent<{ audio?: boolean; video?: boolean }>) => void this.teams.client?.setPrejoin(e.detail)}
					@office-prejoin-join=${() => void this.teams.client?.joinCall()}
					@office-prejoin-cancel=${() => ((this.meeting = false), this.teams.client?.leaveCall())}
				></office-ui-prejoin>
			`;
		const self = call.self;
		return html`
			<section class="meeting">
				<office-ui-call-controls
					heading=${`# ${call.channelName}`}
					started-at=${call.startedAt ?? 0}
					?mic=${self?.state.audio}
					?camera=${self?.state.video}
					?screen=${self?.state.screen}
					?hand=${self?.state.hand}
					?share-unavailable=${typeof navigator.mediaDevices?.getDisplayMedia !== 'function'}
					panel=${this.panel}
					@office-call-action=${(e: CustomEvent<{ action: string }>) => this.onCallAction(e.detail.action)}
				></office-ui-call-controls>
				<div class="stage">
					<office-ui-call-grid
						.participants=${call.participants.map((p) => ({
							id: p.id,
							name: p.name,
							self: p.self,
							stream: p.stream as unknown as MediaStream | null,
							audio: p.state.audio,
							video: p.state.video,
							screen: p.state.screen,
							hand: p.state.hand,
							connection: p.connection,
						}))}
					></office-ui-call-grid>
					${this.panel === 'chat'
						? html`<aside class="panel chat">${this.conversation(s, true)}</aside>`
						: this.panel === 'people'
							? html`<aside class="panel people-panel">
									<h2>People (${call.participants.length})</h2>
									<ul class="plain">
										${call.participants.map(
											(p) => html`<li>
												<office-ui-avatar size="sm" name=${p.name} seed=${p.id}></office-ui-avatar>
												<span>${p.name}${p.self ? ' (you)' : ''}</span>
												${p.state.hand ? html`<span title="Hand raised">✋</span>` : nothing}
											</li>`,
										)}
									</ul>
								</aside>`
							: nothing}
				</div>
			</section>
		`;
	}
}

export function defineTeamsApp(registry: CustomElementRegistry | undefined = globalThis.customElements): void {
	if (registry && !registry.get('teams-app')) registry.define('teams-app', TeamsApp);
}
