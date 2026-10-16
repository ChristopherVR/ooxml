// <teams-app>: one reactive element, one template, one CSS file. All collaboration and call logic
// lives in `createTeamsClient` (ooxml-core/teams); the visual pieces are ooxml-ui elements. This
// file only decides what is on screen (`render`) and forwards user intent to client actions.
import { type LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import { TeamsLitElement } from '../i18n/lit-element';
import { describeActivityIn } from '../i18n/describe';
import { localizeCoreText } from '../i18n/core-text';
import type { TeamsKey } from '../i18n/translator';
import {
	type ActivityFilter,
	type ActivityItem,
	type FileUploader,
	type ChannelTab,
	type OfficeKind,
	type TeamsClient,
	type TeamsServerConfig,
	type TeamsState,
	type DraftContext,
	type MeetingFormInput,
	type TabContent,
	isPresenceStatus,
	isStatusDuration,
	parseServerConfig,
	type PresenceStatus,
	type StatusDuration,
} from 'ooxml-core/teams';
// Registered from the leaf modules, not the package root: the root imports this folder through
// `teams/index`, so importing it back would make `OFFICE_UI_TAGS` read `TEAMS_TAGS` too early.
import { registerControls } from '../../controls';
import { definePresence } from '../../presence';
import { installOfficeUiTheme, THEME_CSS } from '../../theme';
import { registerTeams } from '../index';
import { TeamsController } from './controller';
import { defineTeamsContentPreview, type FileEmbeds, type SaveFileCopy } from './content-preview';
import { defineTeamsChannelTab } from './channel-tab';
import { defineTeamsFilesPanel } from './files-panel';
import { defineTeamsSettings, type TeamsTheme } from './teams-settings';
import {
	loadConfig,
	loadIdentity,
	safeStorage,
	saveConfig,
	saveIdentity,
	type Identity,
} from './storage';
import css from './teams-app.css?raw';
import { threadPane } from './thread-pane';
import { followedThreads } from './followed-threads';
import { draftList } from './draft-list';
import { filePopoutDetail, filePopoutUrl } from './file-popout';
import { messageTransfers } from './message-transfers';
import { defineTeamsProfileMenu } from './profile-menu';
import { SignInFlow } from './sign-in-flow';
import { defineTeamsSignIn } from './teams-sign-in';
import { defineTeamsPinnedBar } from './pinned-bar';
import { watchActivity } from './activity';
import { browserDeliveryEnv, deliverNotification } from './notify-delivery';
import { defineTeamsAddTabDialog } from './add-tab-dialog';
import { icon } from '../icons';
import type { ChatSendDetail } from '../chat-composer';
import { defineTeamsCreateChannelDialog } from './create-channel-dialog';
import { defineTeamsNewChatDialog } from './new-chat-dialog';
import { defineTeamsChatPeopleDialog } from './chat-people-dialog';
import { defineTeamsMeetingDialog } from './meeting-dialog';
import { type MeetingChange, defineTeamsMeetingDetails } from './meeting-details';
import { chatHeader, meetingCards } from './chat-app';
import { downloadText } from './calendar-app';

export type { FileUploader } from 'ooxml-core/teams';
export interface OpenFileDetail {
	/** Channel captured when content opens, so saving cannot target a later selection. */
	channelId?: string;
	attachment: {
		name: string;
		kind: OfficeKind;
		size?: number | undefined;
		url?: string | undefined;
		mime?: string | undefined;
	};
	/** The resolved URL, signed when the file lives on your authenticated server. */
	url: string | undefined;
}
export type FileOpeners = Partial<Record<OfficeKind, (detail: OpenFileDetail) => void>>;

type RailView =
	| 'activity'
	| 'chat'
	| 'teams'
	| 'calendar'
	| 'calls'
	| 'files'
	| 'followed'
	| 'drafts';
type Panel = '' | 'chat' | 'people';
const LOCAL_THEME = THEME_CSS.replaceAll(
	':root[data-office-theme="dark"]',
	':host([data-office-theme="dark"])',
)
	.replaceAll(':root:not([data-office-theme="light"])', ':host([data-office-theme="system"])')
	.replaceAll(':root', ':host([data-office-theme])');
const RAIL = [
	{ id: 'activity', label: 'rail.activity', icon: 'bell' },
	{ id: 'chat', label: 'rail.chat', icon: 'message' },
	{ id: 'teams', label: 'rail.teams', icon: 'users' },
	{ id: 'calendar', label: 'rail.calendar', icon: 'calendar' },
	{ id: 'calls', label: 'rail.calls', icon: 'phone', secondary: true },
	{ id: 'files', label: 'rail.files', icon: 'folder', secondary: true },
	{
		id: 'followed',
		label: 'rail.followed',
		shortLabel: 'rail.threads',
		icon: 'chat',
		secondary: true,
	},
	{ id: 'drafts', label: 'rail.drafts', icon: 'chat', secondary: true },
] as const satisfies readonly {
	id: string;
	label: TeamsKey;
	shortLabel?: TeamsKey;
	icon: string;
	secondary?: boolean;
}[];

export class TeamsApp extends TeamsLitElement {
	static override styles = [unsafeCSS(LOCAL_THEME), unsafeCSS(css)];
	static override properties = {
		workspaceId: { type: String, attribute: 'workspace-id' },
		userName: { type: String, attribute: 'user-name' },
		userId: { type: String, attribute: 'user-id' },
		serverConfig: { type: String, attribute: 'server-config' },
		config: { attribute: false },
		uploadFile: { attribute: false },
		openers: { attribute: false },
		embeds: { attribute: false },
		preview: { state: true },
		addingTab: { state: true },
		creatingChannel: { state: true },
		newChat: { state: true },
		chatPeople: { state: true },
		meetingForm: { state: true },
		meetingDetails: { state: true },
		rail: { state: true },
		tab: { state: true },
		panel: { state: true },
		meeting: { state: true },
		settingsOpen: { state: true },
		mobileView: { state: true },
		toast: { state: true },
		activityToasts: { state: true },
		activityFilter: { state: true },
		highlightId: { state: true },
		identity: { state: true },
		followedUnreadOnly: { state: true },
		theme: { state: true },
		fileOpenPreference: { state: true },
		chatDensity: { state: true },
		showAppNames: { state: true },
		authRevision: { state: true },
	};
	declare workspaceId: string;
	declare userName: string;
	declare userId: string;
	declare serverConfig: string;
	/** Server settings. Omit to use the settings dialog (remembered in this browser) or local mode. */
	declare config: TeamsServerConfig | null;
	/** Where attachments go. Default: your server's `/files` endpoint. */
	declare uploadFile: FileUploader | undefined;
	/** Per-product openers override the built-in preview. */
	declare openers: FileOpeners;
	/** Host embedding pages, for example a PowerPoint viewer accepting a file URL. */
	declare embeds: FileEmbeds;
	declare preview: OpenFileDetail | null;
	declare rail: RailView;
	declare tab: string;
	declare addingTab: boolean;
	declare creatingChannel: boolean;
	declare newChat: boolean;
	declare chatPeople: boolean;
	/** The New / Edit meeting form: the meeting being edited, the prefilled start and, when only one
	 * occurrence of a repeating meeting is edited, its original start. */
	declare meetingForm: { meetingId: string; start: number; occurrence?: number } | null;
	/** The meeting details dialog: the meeting and the occurrence that was opened. */
	declare meetingDetails: { meetingId: string; start: number } | null;
	declare panel: Panel;
	declare meeting: boolean;
	declare settingsOpen: boolean;
	/** Compact shell only: the channel list or the conversation (and other content). */
	declare mobileView: 'list' | 'conversation';
	declare toast: string;
	/** The newest activity item announced while the tab is visible. */
	declare activityToasts: ActivityItem[];
	declare activityFilter: ActivityFilter;
	/** A message briefly highlighted after opening it from Activity or the pinned bar. */
	declare highlightId: string;
	declare identity: Identity | null;
	declare followedUnreadOnly: boolean;
	declare theme: TeamsTheme;
	declare fileOpenPreference: 'teams' | 'browser';
	declare chatDensity: 'comfy' | 'compact';
	declare showAppNames: boolean;
	/** Bumped when the sign-in state changes, so the client restarts with or without a session. */
	declare authRevision: number;

	/** Sign-in for servers that require accounts; a server without them never gates the app. */
	private readonly signIn = new SignInFlow({
		changed: () => this.authRevision++,
		emit: (name, detail) =>
			this.dispatchEvent(new CustomEvent(name, { detail, bubbles: true, composed: true })),
		i18n: () => this.i18n,
	});
	private authKey = '';

	private readonly teams = new TeamsController(
		this,
		(text) => this.notify(localizeCoreText(text, this.i18n)),
		(item) => this.announce(item),
	);
	private toastTimer: ReturnType<typeof setTimeout> | undefined;
	private readonly activityToastTimers = new Map<string, ReturnType<typeof setTimeout>>();
	private highlightTimer: ReturnType<typeof setTimeout> | undefined;
	private stopActivity: (() => void) | undefined;
	private openRequest = 0;
	/** The channel last open in the Teams app, restored when you come back from Chat. */
	private lastChannelId = '';
	private retainedTab: ChannelTab | undefined;
	private themeStorageKey = '';
	private readonly popoutRequested =
		new URL(globalThis.location?.href ?? 'https://workspace.invalid').searchParams.get(
			'openteams-file',
		) === '1';
	private readonly popout = filePopoutDetail(
		globalThis.location?.href ?? 'https://workspace.invalid',
	);

	constructor() {
		super();
		this.workspaceId = 'demo';
		this.userName = '';
		this.userId = '';
		this.serverConfig = '';
		this.config = null;
		this.uploadFile = undefined;
		this.openers = {};
		this.embeds = {};
		this.preview = null;
		this.retainedTab = undefined;
		this.addingTab = false;
		this.creatingChannel = false;
		this.newChat = false;
		this.chatPeople = false;
		this.meetingForm = null;
		this.meetingDetails = null;
		this.rail = 'teams';
		this.tab = 'posts';
		this.panel = '';
		this.meeting = false;
		this.settingsOpen = false;
		this.mobileView = 'conversation';
		this.toast = '';
		this.activityToasts = [];
		this.activityFilter = 'all';
		this.highlightId = '';
		this.identity = null;
		this.followedUnreadOnly = false;
		this.theme = 'system';
		this.fileOpenPreference = 'teams';
		this.chatDensity = 'comfy';
		this.showAppNames = true;
		this.authRevision = 0;
	}

	/** The core client behind this element, for hosts that want the raw actions. */
	get client(): TeamsClient | null {
		return this.teams.client;
	}

	override connectedCallback(): void {
		installOfficeUiTheme();
		registerControls();
		definePresence();
		registerTeams();
		defineTeamsSettings();
		defineTeamsContentPreview();
		defineTeamsChannelTab();
		defineTeamsFilesPanel();
		defineTeamsProfileMenu();
		defineTeamsPinnedBar();
		defineTeamsAddTabDialog();
		defineTeamsCreateChannelDialog();
		defineTeamsNewChatDialog();
		defineTeamsChatPeopleDialog();
		defineTeamsMeetingDialog();
		defineTeamsMeetingDetails();
		defineTeamsSignIn();
		// Input anywhere in the page keeps the person from showing as Away.
		this.stopActivity = watchActivity(this.ownerDocument, () => this.teams.client?.noteActivity());
		super.connectedCallback();
		if (this.signIn.phase === 'checking') this.authRevision++;
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
		if (this.popoutRequested) return;
		const base = this.resolveConfig();
		const waiting = this.signIn.gate(base);
		const authKey = `${waiting}|${this.signIn.session?.token ?? ''}`;
		const restart =
			['workspaceId', 'userName', 'userId', 'serverConfig', 'config', 'identity'].some((k) =>
				changed.has(k as never),
			) ||
			(changed.has('authRevision') && authKey !== this.authKey);
		this.authKey = authKey;
		if (!restart && this.teams.client) return;
		this.preview = null;
		this.tab = 'posts';
		this.addingTab = false;
		this.creatingChannel = false;
		this.newChat = this.chatPeople = false;
		this.meetingForm = this.meetingDetails = null;
		this.settingsOpen = false;
		this.openRequest++;
		// A server that needs accounts: nothing starts until someone signs in, and the account's id
		// (not a claimed one) is the user.
		if (waiting) return this.teams.stop();
		const account = this.signIn.session;
		const remembered = this.identity ?? loadIdentity();
		const named = account
			? {
					id: account.userId,
					name:
						this.userName || (remembered?.id === account.userId ? remembered.name : account.userId),
				}
			: this.userName
				? { id: this.userId || loadIdentity()?.id || crypto.randomUUID(), name: this.userName }
				: remembered;
		if (this.userName && !this.userId && !account) saveIdentity(named!);
		if (!named) return this.teams.stop();
		const workspaceId = /^[\w-]{1,100}$/u.test(this.workspaceId) ? this.workspaceId : 'demo';
		this.themeStorageKey = `teams:theme:${encodeURIComponent(workspaceId)}:${encodeURIComponent(named.id)}`;
		const rememberedTheme = safeStorage.getItem(this.themeStorageKey);
		if (!this.hasAttribute('host-theme'))
			this.theme =
				rememberedTheme === 'dark' || rememberedTheme === 'light' ? rememberedTheme : 'system';
		this.applyTheme();
		this.showAppNames =
			safeStorage.getItem(this.themeStorageKey.replace('teams:theme:', 'teams:app-names:')) !==
			'false';
		this.chatDensity =
			safeStorage.getItem(this.themeStorageKey.replace('teams:theme:', 'teams:density:')) ===
			'compact'
				? 'compact'
				: 'comfy';
		this.fileOpenPreference =
			safeStorage.getItem(this.themeStorageKey.replace('teams:theme:', 'teams:file-open:')) ===
			'browser'
				? 'browser'
				: 'teams';
		this.teams.start({
			workspaceId,
			user: named,
			config: this.signIn.apply(base),
			storage: safeStorage,
			...(this.uploadFile ? { uploadFile: this.uploadFile } : {}),
		});
		this.dispatchEvent(
			new CustomEvent('teams-ready', { detail: { user: named }, bubbles: true, composed: true }),
		);
	}

	protected override updated(): void {
		this.setAttribute('data-chat-density', this.chatDensity);
		const state = this.teams.state;
		this.signIn.watchConnection(state?.status === 'connected', this.resolveConfig());
		if (state?.channels.some((c) => c.id === state.selectedChannelId))
			this.lastChannelId = state.selectedChannelId;
		if (
			state?.thread &&
			document.visibilityState === 'visible' &&
			this.renderRoot.querySelector('.thread-pane')
		)
			this.teams.client?.markThreadRead(state.selectedChannelId, state.thread.root.id);
	}

	private notify(text: string): void {
		this.toast = text;
		clearTimeout(this.toastTimer);
		this.toastTimer = setTimeout(() => (this.toast = ''), 4000);
	}

	/** Deliver a notification the core allowed: a toast in a visible tab, else a desktop one. */
	private announce(item: ActivityItem): void {
		const settings = this.teams.client?.getState().notificationSettings;
		if (!settings) return;
		deliverNotification(item, settings, browserDeliveryEnv(), {
			toast: (entry) => {
				// Newest first, at most three on screen, each leaves on its own timer.
				this.activityToasts = [
					entry,
					...this.activityToasts.filter((t) => t.id !== entry.id),
				].slice(0, 3);
				clearTimeout(this.activityToastTimers.get(entry.id));
				this.activityToastTimers.set(
					entry.id,
					setTimeout(() => this.dismissToast(entry.id), 6000),
				);
			},
			open: (entry) =>
				entry.kind === 'meeting'
					? this.openMeeting(entry.messageId)
					: void this.openActivity(entry.id),
		});
	}

	// ---- intent -> client actions -----------------------------------------------------------
	private async openFile(
		a: OpenFileDetail['attachment'],
		mode?: 'teams' | 'browser',
	): Promise<void> {
		if (!this.canLeaveContent()) return;
		const browser =
			mode === 'browser' ||
			(!mode && a.kind !== 'other' && a.kind !== 'vsdx' && this.fileOpenPreference === 'browser');
		const popout = browser ? globalThis.open('about:blank', '_blank') : null;
		if (popout) popout.opener = null;
		const request = ++this.openRequest;
		const channelId =
			(a as { channelId?: string }).channelId ?? this.teams.state?.selectedChannelId;
		const url = await this.teams.client?.fileUrl(a);
		if (request !== this.openRequest || !this.isConnected) {
			popout?.close();
			return;
		}
		if (!this.canLeaveContent()) {
			popout?.close();
			return;
		}
		const detail: OpenFileDetail = { attachment: a, url, ...(channelId ? { channelId } : {}) };
		if (
			!this.dispatchEvent(
				new CustomEvent('teams-open-file', {
					detail,
					bubbles: true,
					composed: true,
					cancelable: true,
				}),
			)
		) {
			popout?.close();
			return;
		}
		const opener = this.openers[a.kind];
		if (opener) {
			popout?.close();
			return opener(detail);
		}
		if (browser && url) {
			const target = filePopoutUrl(detail, this.ownerDocument.location.href);
			if (popout && target) {
				popout.location.replace(target);
				return;
			}
			popout?.close();
			this.notify(this.t('app.allowPopups'));
			return;
		}
		popout?.close();
		if (url) this.preview = detail;
		else this.notify(this.t('app.nameOnly'));
	}

	/** Preview a host-provided website or file without changing shared conversation state. */
	previewContent(detail: OpenFileDetail): void {
		if (!this.canLeaveContent()) return;
		this.openRequest++;
		this.preview = detail;
	}

	private canLeaveContent(): boolean {
		const direct = this.shadowRoot?.querySelector<HTMLElement & { canLeave(): boolean }>(
			'teams-content-preview',
		);
		const nested = this.shadowRoot
			?.querySelector('teams-channel-tab')
			?.shadowRoot?.querySelector<HTMLElement & { canLeave(): boolean }>('teams-content-preview');
		return (direct ?? nested)?.canLeave() ?? true;
	}

	private closePreview(): boolean {
		if (!this.canLeaveContent()) return false;
		this.openRequest++;
		this.preview = null;
		return true;
	}

	override disconnectedCallback(): void {
		this.openRequest++;
		this.preview = null;
		clearTimeout(this.toastTimer);
		for (const timer of this.activityToastTimers.values()) clearTimeout(timer);
		this.activityToastTimers.clear();
		clearTimeout(this.highlightTimer);
		this.stopActivity?.();
		this.stopActivity = undefined;
		this.signIn.dispose();
		super.disconnectedCallback();
	}

	private createChannel(): void {
		this.creatingChannel = true;
	}

	private commitChannel(name: string, description: string, client: TeamsClient | null): boolean {
		if (!this.creatingChannel || !client || client !== this.client || !this.closePreview())
			return false;
		client.createChannel(name, description);
		this.creatingChannel = false;
		this.tab = 'posts';
		this.rail = 'teams';
		this.meeting = false;
		this.mobileView = 'conversation';
		void this.focusCreatedChannel();
		return true;
	}
	private async focusCreatedChannel(): Promise<void> {
		await this.updateComplete;
		await this.renderRoot.querySelector<LitElement>('teams-create-channel-dialog')?.updateComplete;
		const heading = this.renderRoot.querySelector<HTMLElement>('main h1');
		if (heading && !this.creatingChannel) {
			heading.tabIndex = -1;
			heading.focus();
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
		if (!this.closePreview()) return;
		void this.teams.client?.openCall(channelId);
		this.mobileView = 'conversation';
		this.meeting = true;
		const target = channelId ?? this.teams.state?.selectedChannelId;
		this.rail = this.teams.state?.chats.some((c) => c.id === target) ? 'chat' : 'teams';
	}

	private selectChannel(id: string): boolean {
		if (!this.closePreview()) return false;
		this.settingsOpen = false;
		this.addingTab = false;
		this.creatingChannel = false;
		this.teams.client?.select(id);
		this.teams.client?.search('');
		this.meeting = false;
		this.tab = 'posts';
		this.rail = this.teams.state?.chats.some((c) => c.id === id) ? 'chat' : 'teams';
		this.mobileView = 'conversation';
		return true;
	}

	private askSettings(): void {
		this.settingsOpen = true;
	}

	private async closeSettings(): Promise<void> {
		this.settingsOpen = false;
		await this.updateComplete;
		if (!this.isConnected || this.settingsOpen) return;
		this.renderRoot
			.querySelector('office-ui-menu-button.settings-menu')
			?.shadowRoot?.querySelector<HTMLButtonElement>('button')
			?.focus();
	}

	private applyTheme(): void {
		if (this.theme === 'system') this.removeAttribute('data-office-theme');
		else this.setAttribute('data-office-theme', this.theme);
	}

	/** Which screen the compact shell shows: the channel or call list, or the content itself. */
	private mobileScreen(s: TeamsState): 'list' | 'content' {
		const listed = ['teams', 'calls', 'activity', 'chat', 'calendar'].includes(this.rail);
		return listed && this.mobileView === 'list' && !(this.meeting && s.call) ? 'list' : 'content';
	}

	// ---- template ---------------------------------------------------------------------------
	protected override render() {
		if (this.popoutRequested)
			return this.popout
				? html`<div class="popout-view">
						<teams-content-preview
							.detail=${this.popout}
							@teams-preview-close=${() => globalThis.close()}
						></teams-content-preview>
					</div>`
				: html`<p role="alert">${this.t('app.invalidPreview')}</p>`;
		if (this.signIn.phase !== 'open') return this.signInScreen();
		const s = this.teams.state;
		if (!s) return this.welcome();
		return html`
			<div
				class="shell"
				data-mobile-view=${this.mobileScreen(s)}
				@office-chat-open-tab=${(event: CustomEvent<{ tabId: string }>) => {
					if (this.teams.state?.tabs.some((entry) => entry.id === event.detail.tabId))
						this.selectTab(event.detail.tabId);
					else this.notify(this.t('app.tabGone'));
				}}
			>
				${this.topbar(s)}
				<div class="body">
					<office-ui-app-rail
						.hideLabels=${!this.showAppNames}
						.items=${RAIL.map((i) => ({ ...i, label: this.t(i.label), ...('shortLabel' in i ? { shortLabel: this.t(i.shortLabel) } : {}), ...(i.id === 'calls' && s.channels.some((c) => c.live) ? { badge: s.channels.filter((c) => c.live).length } : {}), ...(i.id === 'activity' && s.activityUnread ? { badge: s.activityUnread } : {}), ...(i.id === 'chat' && s.chats.some((c) => c.unread) ? { badge: s.chats.reduce((n, c) => n + c.unread, 0) } : {}) }))}
						selected=${this.rail}
						@office-rail-select=${(e: CustomEvent<{ id: RailView }>) => {
							if (!this.closePreview()) return;
							this.settingsOpen = false;
							if (e.detail.id === 'teams') this.leaveChatSelection();
							this.rail = e.detail.id;
							this.meeting = false;
							this.mobileView = 'list';
						}}
					></office-ui-app-rail>
					<aside class="side" ?inert=${this.settingsOpen}>${this.sidebar(s)}</aside>
					<main class="main" ?inert=${this.settingsOpen}>${this.main(s)}</main>
					<teams-settings
						embedded
						?open=${this.settingsOpen}
						.config=${this.resolveConfig()}
						.theme=${this.theme}
						.fileOpenPreference=${this.fileOpenPreference}
						.chatDensity=${this.chatDensity}
						.showAppNames=${this.showAppNames}
						@teams-settings-app-names=${(event: CustomEvent<{ show: boolean }>) => {
							if (typeof event.detail.show !== 'boolean') return;
							this.showAppNames = event.detail.show;
							safeStorage.setItem(
								this.themeStorageKey.replace('teams:theme:', 'teams:app-names:'),
								String(this.showAppNames),
							);
						}}
						@teams-settings-density=${(event: CustomEvent<{ density: 'comfy' | 'compact' }>) => {
							if (!['comfy', 'compact'].includes(event.detail.density)) return;
							this.chatDensity = event.detail.density;
							safeStorage.setItem(
								this.themeStorageKey.replace('teams:theme:', 'teams:density:'),
								this.chatDensity,
							);
						}}
						@teams-settings-file-open=${(
							event: CustomEvent<{ preference: 'teams' | 'browser' }>,
						) => {
							if (!['teams', 'browser'].includes(event.detail.preference)) return;
							this.fileOpenPreference = event.detail.preference;
							safeStorage.setItem(
								this.themeStorageKey.replace('teams:theme:', 'teams:file-open:'),
								this.fileOpenPreference,
							);
						}}
						.followSettings=${s.threadFollowSettings}
						.notificationSettings=${s.notificationSettings}
						@teams-settings-notifications=${(event: CustomEvent<Partial<typeof s.notificationSettings>>) => this.teams.client?.setNotificationSettings(event.detail)}
						.userName=${s.user.name}
						@teams-settings-theme=${(event: CustomEvent<{ theme: TeamsTheme }>) => {
							if (!['system', 'light', 'dark'].includes(event.detail.theme)) return;
							this.theme = event.detail.theme;
							this.applyTheme();
							safeStorage.setItem(this.themeStorageKey, this.theme);
						}}
						@teams-settings-follow=${(event: CustomEvent<Partial<typeof s.threadFollowSettings>>) => this.teams.client?.setThreadFollowSettings(event.detail)}
						@teams-settings-close=${() => void this.closeSettings()}
						@teams-settings-apply=${(e: CustomEvent<{ config: TeamsServerConfig }>) => {
							if (!this.closePreview()) return;
							saveConfig(e.detail.config);
							this.config = e.detail.config;
							this.settingsOpen = false;
							this.dispatchEvent(
								new CustomEvent('teams-config-change', {
									detail: e.detail,
									bubbles: true,
									composed: true,
								}),
							);
						}}
					></teams-settings>
				</div>
				<teams-create-channel-dialog
					?open=${this.creatingChannel}
					.create=${(name: string, description: string) => this.commitChannel(name, description, this.client)}
					@teams-create-channel-close=${() => (this.creatingChannel = false)}
				></teams-create-channel-dialog>
				${this.toast ? html`<div class="toast" role="status">${this.toast}</div>` : nothing}
				${this.activityToasts.length ? html`<div class="activity-toasts">${this.activityToasts.map((t) => this.activityToastView(t))}</div>` : nothing}
				${this.chatDialogs(s)} ${this.calendarDialogs(s)}
				<teams-add-tab-dialog
					.open=${this.addingTab}
					.client=${this.teams.client}
					.channelId=${s.selectedChannelId}
					.channelName=${s.channel?.name ?? ''}
					.files=${s.files}
					.add=${(
						name: string,
						content: TabContent,
						channelId: string,
						client: TeamsClient,
						postToChannel: boolean,
					) => {
						if (client !== this.teams.client || client.getState().selectedChannelId !== channelId)
							return null;
						if (!this.closePreview()) return 'canceled';
						const tab = client.addTab(name, content, { postToChannel });
						if (tab) {
							this.retainedTab = tab;
							this.tab = tab.id;
							this.addingTab = false;
							this.creatingChannel = false;
						}
						return tab;
					}}
					@teams-add-tab-close=${() => (this.addingTab = false)}
				></teams-add-tab-dialog>
			</div>
		`;
	}

	private signInScreen() {
		const flow = this.signIn;
		return html`<teams-sign-in
			.userId=${flow.userId}
			.notice=${flow.notice}
			.error=${flow.error}
			.retryAfter=${flow.retryAfterSec}
			?busy=${flow.busy}
			?checking=${flow.phase === 'checking'}
			@teams-sign-in-submit=${(event: CustomEvent<{ userId: string; credential: string }>) =>
				void flow.submit(this.resolveConfig(), event.detail.userId, event.detail.credential)}
		></teams-sign-in>`;
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
					<h1>${this.t('welcome.title')}</h1>
					<p>${this.t('welcome.pickName')}</p>
					<input
						name="name"
						placeholder=${this.t('welcome.yourName')}
						aria-label=${this.t('welcome.yourName')}
						maxlength="40"
						autocomplete="name"
						required
					/>
					<button type="submit">${this.t('welcome.continue')}</button>
				</form>
			</div>
		`;
	}

	private topbar(s: TeamsState) {
		return html`
			<header class="topbar">
				<div class="brand"><span class="logo">T</span><span>${this.t('rail.teams')}</span></div>
				<div class="search">
					<input
						type="search"
						placeholder=${this.t('search.messages')}
						aria-label=${this.t('search.messages')}
						.value=${s.searchQuery}
						@input=${(e: Event) => this.teams.client?.search((e.target as HTMLInputElement).value)}
					/>
					${
						s.searchQuery
							? html`<ul class="results" role="listbox" aria-label=${this.t('search.results')}>
									${
										s.searchResults.length === 0
											? html`<li class="none">${this.t('search.none')}</li>`
											: s.searchResults.map(
													(h) =>
														html`<li>
															<button
																type="button"
																@click=${() => {
																	if (this.selectChannel(h.channelId)) {
																		this.teams.client?.search('');
																		void this.openThread(h.message.id);
																	}
																}}
															>
																<strong>${h.message.authorName}</strong>
																<span class="in"
																	>${this.t('search.in', { where: `${s.chats.some((c) => c.id === h.channelId) ? '' : '# '}${h.channelName}` })}</span
																>
																<span class="snippet">${h.message.text.slice(0, 100)}</span>
															</button>
														</li>`,
												)
									}
								</ul>`
							: nothing
					}
				</div>
				<div class="me">
					<office-ui-menu-button
						class="settings-menu"
						label=${this.t('app.settingsAndMore')}
						icon="more"
						icon-only
						@office-command=${(event: CustomEvent<{ command: string }>) => {
							if (event.detail.command === 'teams-settings') this.askSettings();
						}}
					>
						<office-ui-menu-item
							label=${this.t('settings.title')}
							command="teams-settings"
						></office-ui-menu-item>
					</office-ui-menu-button>
					<teams-profile-menu
						.state=${s}
						.account=${this.signIn.session?.userId ?? ''}
						@teams-sign-out=${() => void this.signIn.signOut(this.resolveConfig())}
						@teams-profile-status=${(
							event: CustomEvent<{ availability: PresenceStatus; duration?: StatusDuration }>,
						) => {
							const { availability, duration } = event.detail;
							if (isPresenceStatus(availability))
								this.teams.client?.setAvailability(
									availability,
									isStatusDuration(duration) ? duration : 'never',
								);
						}}
						@teams-profile-message=${(
							event: CustomEvent<{ message: string; duration?: StatusDuration }>,
						) =>
							this.teams.client?.setStatusMessage(
								String(event.detail.message ?? ''),
								isStatusDuration(event.detail.duration) ? event.detail.duration : 'never',
							)}
						@teams-profile-idle=${(event: CustomEvent<{ minutes: number }>) =>
							this.teams.client?.setIdleMinutes(Number(event.detail.minutes))}
					></teams-profile-menu>
				</div>
			</header>
		`;
	}

	private dismissToast(id: string): void {
		clearTimeout(this.activityToastTimers.get(id));
		this.activityToastTimers.delete(id);
		this.activityToasts = this.activityToasts.filter((t) => t.id !== id);
	}

	private activityToastView(item: ActivityItem) {
		return html`<div class="activity-toast" role="status" data-kind=${item.kind}>
			<span class="toast-text"
				><strong>${describeActivityIn(item, this.i18n)}</strong
				><span class="toast-preview">${item.preview}</span></span
			>
			<button
				type="button"
				@click=${() => {
					this.dismissToast(item.id);
					if (item.kind === 'meeting') this.openMeeting(item.messageId);
					else void this.openActivity(item.id);
				}}
			>
				${this.t('app.view')}
			</button>
			<button
				type="button"
				aria-label=${this.t('app.dismiss')}
				@click=${() => this.dismissToast(item.id)}
			>
				${icon('close')}
			</button>
		</div>`;
	}

	private sidebar(s: TeamsState) {
		if (this.rail === 'chat')
			return html`<office-ui-chat-sidebar
				.chats=${s.chats}
				selected-id=${s.selectedChannelId}
				@office-chat-select=${(e: CustomEvent<{ id: string }>) => this.selectChannel(e.detail.id)}
				@office-chat-new=${() => (this.newChat = true)}
			></office-ui-chat-sidebar>`;
		if (this.rail === 'calendar')
			return html`<office-ui-calendar-agenda
					.meetings=${s.meetings}
					.zone=${s.viewerZone}
					@office-agenda-open=${(e: CustomEvent<{ meetingId: string; start: number }>) => (this.meetingDetails = { meetingId: e.detail.meetingId, start: e.detail.start })}
					@office-agenda-join=${(e: CustomEvent<{ meetingId: string }>) => void this.joinMeeting(e.detail.meetingId)}
					@office-agenda-new=${() => (this.meetingForm = { meetingId: '', start: 0 })}
				></office-ui-calendar-agenda>
				<button
					type="button"
					class="show-calendar"
					@click=${() => (this.mobileView = 'conversation')}
				>
					${icon('calendar')} ${this.t('app.openCalendar')}
				</button>`;
		if (this.rail === 'activity')
			return html`<office-ui-activity-feed
				.items=${s.activity}
				.filter=${this.activityFilter}
				@office-activity-filter=${(e: CustomEvent<{ filter: ActivityFilter }>) => (this.activityFilter = e.detail.filter)}
				@office-activity-open=${(e: CustomEvent<{ id: string }>) => void this.openActivity(e.detail.id)}
				@office-activity-read=${(e: CustomEvent<{ id: string; read: boolean }>) => this.teams.client?.markActivityRead(e.detail.id, e.detail.read)}
				@office-activity-read-all=${() => this.teams.client?.markAllActivityRead()}
			></office-ui-activity-feed>`;
		if (this.rail === 'calls') {
			const live = s.channels.filter((c) => c.live);
			return html`
				<h2>${this.t('rail.calls')}</h2>
				<ul class="plain">
					${s.channels.map(
						(c) =>
							html`<li>
								<span># ${c.name}</span>
								<button
									type="button"
									class=${c.live ? 'join live' : 'join'}
									@click=${() => this.startMeeting(c.id)}
								>
									${this.t(c.live ? 'call.join' : 'call.meetNow')}
								</button>
							</li>`,
					)}
				</ul>
				${live.length === 0 ? html`<p class="hint">${this.t('app.nobodyInCall')}</p>` : nothing}
			`;
		}
		return html`
			<office-ui-channel-list
				selected-id=${s.selectedChannelId}
				.channels=${s.channels}
				@office-channel-select=${(e: CustomEvent<{ id: string }>) => this.selectChannel(e.detail.id)}
				@office-channel-create=${this.createChannel}
			></office-ui-channel-list>
		`;
	}

	private main(s: TeamsState) {
		if (this.rail === 'drafts' && this.teams.client)
			return draftList(
				s,
				this.teams.client,
				(context) => void this.resumeDraft(context),
				this.i18n,
			);
		if (this.preview)
			return html`<teams-content-preview
				.detail=${this.preview}
				.embeds=${this.embeds}
				.saveCopy=${this.saveCopyFor(this.preview, s)}
				@teams-preview-close=${this.closePreview}
			></teams-content-preview>`;
		if (this.rail === 'files') return this.allFiles(s);
		if (this.rail === 'followed' && this.teams.client)
			return followedThreads(
				s,
				this.teams.client,
				(thread) => {
					if (this.selectChannel(thread.channelId)) void this.openThread(thread.root.id);
				},
				this.followedUnreadOnly,
				(value) => {
					this.followedUnreadOnly = value;
				},
				this.i18n,
			);
		if (s.call && this.meeting) return this.meetingView(s);
		if (this.rail === 'calendar') return this.calendarMain(s);
		if (this.rail === 'chat' && !s.chats.some((c) => c.id === s.selectedChannelId))
			return this.chatEmpty();
		return html`
			${
				s.call && !this.meeting
					? html`<div class="callbar">
							${this.t('app.inCall', { where: this.callHeading(s.call) })}
							<button
								type="button"
								@click=${() => {
									if (s.call!.kind === 'meeting') {
										if (this.closePreview()) [this.rail, this.meeting] = ['calendar', true];
									} else if (this.selectChannel(s.call!.channelId)) this.meeting = true;
								}}
							>
								${this.t('app.returnToCall')}
							</button>
						</div>`
					: nothing
			}
			${this.conversation(s)}
		`;
	}

	private conversation(s: TeamsState, compact = false) {
		const channel = s.channel;
		const chat = s.chats.find((x) => x.id === s.selectedChannelId);
		const c = this.teams.client;
		const compose = s.editing ?? s.replyingTo;
		const sharedTab = s.tabs.find((t) => t.id === this.tab);
		if (sharedTab) this.retainedTab = sharedTab;
		const activeTab =
			sharedTab ??
			(this.retainedTab?.id === this.tab && this.retainedTab.channelId === s.selectedChannelId
				? this.retainedTab
				: undefined);
		return html`
			${
				compact
					? nothing
					: chat
						? chatHeader(
								chat,
								{
									back: () => (this.mobileView = 'list'),
									call: () => this.startMeeting(),
									people: () => (this.chatPeople = true),
								},
								this.i18n,
							)
						: html`<header class="channel-head">
								<button
									type="button"
									class="back"
									aria-label=${this.t('app.backToChannels')}
									title=${this.t('app.backToChannels')}
									@click=${() => (this.mobileView = 'list')}
								>
									${icon('back')}
								</button>
								<h1><span class="hash">#</span> ${channel?.name ?? this.t('app.noChannel')}</h1>
								<nav class="tabs" role="tablist">
									${(['posts', 'files'] as const).map(
										(t) =>
											html`<button
												type="button"
												role="tab"
												aria-selected=${String(this.tab === t)}
												@click=${() => this.selectTab(t)}
											>
												${this.t(t === 'posts' ? 'app.posts' : 'app.shared')}
											</button>`,
									)}
									${s.tabs.map((t) => html`<button type="button" role="tab" aria-selected=${String(this.tab === t.id)} @click=${() => this.selectTab(t.id)}>${t.name}</button>`)}
								</nav>
								<button
									type="button"
									aria-label=${this.t('app.addTab')}
									title=${this.t('app.addTab')}
									?disabled=${!channel}
									@click=${() => (this.addingTab = !this.addingTab)}
								>
									+
								</button>
								${
									sharedTab
										? html`<button
												type="button"
												data-tab-conversation
												aria-label=${this.t('app.tabConversation')}
												title=${this.t('app.tabConversation')}
												aria-expanded=${String(s.thread?.root.tabId === sharedTab.id)}
												@click=${async () => {
													if (s.thread?.root.tabId === sharedTab.id) {
														await this.closeThread();
														return;
													}
													if (!c?.openTabConversation(sharedTab.id)) {
														this.notify(this.t('app.tabConversationFailed'));
														return;
													}
													await Promise.resolve();
													await this.updateComplete;
													this.shadowRoot
														?.querySelector<HTMLElement>('[data-thread-heading]')
														?.focus();
												}}
											>
												${icon('chat')}
											</button>`
										: nothing
								}
								${
									sharedTab?.createdBy === s.user.id
										? html`<office-ui-menu-button
												label=${this.t('app.tabOptions')}
												icon="more"
												icon-only
												@office-command=${(event: CustomEvent<{ command: string }>) => {
													if (event.detail.command === 'rename-tab') {
														const name = globalThis
															.prompt?.(this.t('app.tabNamePrompt'), sharedTab.name)
															?.trim();
														if (name) c?.renameTab(sharedTab.id, name);
													} else if (event.detail.command === 'remove-tab') {
														if (!this.canLeaveContent()) return;
														if (
															globalThis.confirm?.(
																this.t('app.confirmRemoveTab', { name: sharedTab.name }),
															) === true &&
															c?.removeTab(sharedTab.id)
														) {
															this.openRequest++;
															this.preview = null;
															this.tab = 'posts';
														}
													}
												}}
												><office-ui-menu-item
													label=${this.t('app.renameTab')}
													command="rename-tab"
												></office-ui-menu-item
												><office-ui-menu-item
													label=${this.t('app.removeTab')}
													command="remove-tab"
												></office-ui-menu-item
											></office-ui-menu-button>`
										: nothing
								}
								<span class="topic">${channel?.topic ?? ''}</span>
								<div class="people" aria-label=${this.t('app.peopleOnline')}>
									${s.people.slice(0, 5).map((p) => html`<office-ui-avatar size="sm" name=${p.name} seed=${p.id} color=${p.color} presence=${p.availability}></office-ui-avatar>`)}
									${s.people.length > 5 ? html`<span class="more">+${s.people.length - 5}</span>` : nothing}
								</div>
								<button
									type="button"
									class="meet"
									?disabled=${!channel}
									@click=${() => this.startMeeting()}
								>
									${this.t(s.channels.find((x) => x.id === s.selectedChannelId)?.live ? 'call.joinCall' : 'call.meetNow')}
								</button>
							</header>`
			}
			${c ? messageTransfers(s, c, this.i18n) : nothing}
			${activeTab && !sharedTab && !compact ? html`<p role="status">${this.t('app.tabRemoved')}</p>` : nothing}
			${
				activeTab && !compact
					? html`<div
							class="conversation-grid tab-conversation"
							data-thread=${String(s.thread?.root.tabId === activeTab.id)}
						>
							<div class="conversation-main">
								<teams-channel-tab
									.tab=${activeTab}
									.client=${c ?? null}
									.embeds=${this.embeds}
									.canSave=${s.canUploadFiles}
									@teams-preview-close=${() => this.selectTab('posts')}
								></teams-channel-tab>
							</div>
							${
								s.thread?.root.tabId === activeTab.id && c
									? threadPane(
											s,
											c,
											(attachment) => void this.openFile(attachment),
											() => void this.closeThread(),
											this.highlightId,
											this.i18n,
										)
									: nothing
							}
						</div>`
					: this.tab === 'files' && !compact
						? this.fileList(s.files, s)
						: html`<div class="conversation-grid" data-thread=${String(!!s.thread && !compact)}>
								<div class="conversation-main">
									${
										compact || s.pinned.length === 0
											? nothing
											: html`<teams-pinned-bar
													.pinned=${s.pinned}
													@teams-pinned-open=${(e: CustomEvent<{ messageId: string }>) => void this.revealMessage(e.detail.messageId)}
													@teams-pinned-unpin=${(e: CustomEvent<{ messageId: string }>) => c?.unpinMessage(e.detail.messageId)}
												></teams-pinned-bar>`
									}
									<office-ui-chat-list
										.messages=${compact ? s.messages : s.posts}
										.replyCounts=${compact ? {} : s.replyCounts}
										.pinnedIds=${s.pinned.map((p) => p.message.id)}
										self-id=${s.user.id}
										.highlightId=${this.highlightId}
										.meetings=${meetingCards(s.meetings, Date.now())}
										@office-chat-meeting=${(e: CustomEvent<{ meetingId: string; action: 'join' | 'open' }>) => (e.detail.action === 'join' ? void this.joinMeeting(e.detail.meetingId) : this.openMeeting(e.detail.meetingId))}
										@office-chat-pin=${(e: CustomEvent<{ messageId: string }>) => c?.pinMessage(e.detail.messageId)}
										@office-chat-unpin=${(e: CustomEvent<{ messageId: string }>) => c?.unpinMessage(e.detail.messageId)}
										@office-chat-react=${(e: CustomEvent<{ messageId: string; emoji: string }>) => c?.toggleReaction(e.detail.messageId, e.detail.emoji)}
										@office-chat-thread=${(e: CustomEvent<{ messageId: string }>) => void this.openThread(e.detail.messageId)}
										@office-chat-reply=${(e: CustomEvent<{ messageId: string }>) => (compact ? c?.startReply(e.detail.messageId) : void this.openThread(e.detail.messageId, 'reply'))}
										@office-chat-edit=${(e: CustomEvent<{ messageId: string }>) => (compact ? c?.startEdit(e.detail.messageId) : void this.openThread(e.detail.messageId, 'edit'))}
										@office-chat-delete=${(e: CustomEvent<{ messageId: string }>) => globalThis.confirm?.(this.t('chat.confirmDelete')) !== false && c?.deleteMessage(e.detail.messageId)}
										@office-chat-open-file=${(e: CustomEvent<{ attachment: OpenFileDetail['attachment'] & { kind: OfficeKind } }>) => this.openFile(e.detail.attachment)}
									></office-ui-chat-list>
									${
										s.thread && !compact
											? nothing
											: html`<office-ui-chat-composer
													.typing=${s.typing}
													.replyingTo=${s.replyingTo?.authorName ?? null}
													.editing=${s.editing !== null}
													.value=${s.draft.text}
													.files=${s.draft.files}
													.missingFiles=${s.draft.missingFiles}
													.mentionTargets=${s.mentionTargets}
													@office-chat-draft=${(e: CustomEvent<TeamsState['draft']>) => c?.setDraft(e.detail)}
													?disabled=${!channel}
													placeholder=${chat ? this.t('app.messageTo', { name: chat.title }) : channel ? this.t('app.messageTo', { name: `# ${channel.name}` }) : this.t('app.createChannelFirst')}
													@office-chat-typing=${() => c?.notifyTyping()}
													@office-chat-cancel=${() => c?.cancelCompose()}
													@office-chat-send=${(e: CustomEvent<ChatSendDetail>) => void c?.send(e.detail)}
													data-compose=${compose ? 'on' : 'off'}
												></office-ui-chat-composer>`
									}
								</div>
								${
									!compact && c
										? threadPane(
												s,
												c,
												(attachment) => void this.openFile(attachment),
												() => void this.closeThread(),
												this.highlightId,
												this.i18n,
											)
										: nothing
								}
							</div>`
			}
		`;
	}

	private selectTab(id: string): void {
		if (!this.closePreview()) return;
		this.teams.client?.closeThread();
		this.tab = id;
		this.addingTab = false;
		this.creatingChannel = false;
	}

	/** Scroll a message into view and highlight it, opening its thread when it is a reply. */
	private async revealMessage(id: string): Promise<void> {
		const message = this.teams.state?.messages.find((m) => m.id === id);
		if (!message) return;
		if (message.replyTo) await this.openThread(id);
		await this.updateComplete;
		const lists = [...this.renderRoot.querySelectorAll<LitElement>('office-ui-chat-list')];
		await Promise.all(lists.map((list) => list.updateComplete));
		const row = lists
			.flatMap((list) => [
				...(list.shadowRoot?.querySelectorAll<HTMLElement>('[data-message-id]') ?? []),
			])
			.find((element) => element.dataset.messageId === id);
		row?.scrollIntoView({ block: 'center' });
		row?.focus();
		this.highlightId = id;
		clearTimeout(this.highlightTimer);
		this.highlightTimer = setTimeout(() => (this.highlightId = ''), 2500);
	}

	/** Open an Activity item: its channel or thread, scrolled to the message, and mark it read. */
	private async openActivity(id: string): Promise<void> {
		const c = this.teams.client;
		const item = c?.getState().activity.find((entry) => entry.id === id);
		if (!c || !item || !this.closePreview()) return;
		c.markActivityRead(id);
		this.settingsOpen = false;
		this.addingTab = false;
		this.creatingChannel = false;
		this.meeting = false;
		this.tab = 'posts';
		this.mobileView = 'conversation';
		c.select(item.channelId);
		c.search('');
		if (item.rootId === item.messageId) c.closeThread();
		await new Promise((resolve) => setTimeout(resolve, 0));
		await this.updateComplete;
		await this.revealMessage(item.messageId);
	}

	private async openThread(id: string, action?: 'reply' | 'edit'): Promise<void> {
		const client = this.teams.client;
		client?.openThread(id);
		if (action === 'reply') client?.startReply(id);
		if (action === 'edit') client?.startEdit(id);
		await Promise.resolve();
		await this.updateComplete;
		this.shadowRoot?.querySelector<HTMLElement>('[data-thread-heading]')?.focus();
	}
	private callHeading(call: NonNullable<TeamsState['call']>): string {
		return call.kind === 'channel' ? `# ${call.channelName}` : call.channelName;
	}

	/** Teams rail: if a chat is selected, go back to the channel you were in. */
	private leaveChatSelection(): void {
		const st = this.teams.state;
		const c = this.teams.client;
		if (!st || !c || !st.chats.some((x) => x.id === st.selectedChannelId)) return;
		const target = st.channels.some((x) => x.id === this.lastChannelId)
			? this.lastChannelId
			: st.channels[0]?.id;
		if (target) c.select(target);
	}

	private chatEmpty() {
		return html`<div class="empty-main">
			<h1>${this.t('rail.chat')}</h1>
			<p>${this.t('app.pickChat')}</p>
			<button type="button" @click=${() => (this.newChat = true)}>${this.t('chat.new')}</button>
		</div>`;
	}

	private commitChat(ids: string[], name: string): boolean {
		const c = this.client;
		if (!c || !this.closePreview()) return false;
		const id = c.createGroupChat(ids, name);
		if (!id) return false;
		this.newChat = false;
		this.rail = 'chat';
		this.meeting = false;
		this.tab = 'posts';
		this.mobileView = 'conversation';
		return true;
	}

	private chatDialogs(s: TeamsState) {
		const c = this.teams.client;
		const chat = s.chats.find((x) => x.id === s.selectedChannelId) ?? null;
		return html`
			<teams-new-chat-dialog
				?open=${this.newChat}
				.contacts=${s.contacts}
				.create=${(ids: string[], name: string) => this.commitChat(ids, name)}
				@teams-dialog-close=${() => (this.newChat = false)}
			></teams-new-chat-dialog>
			<teams-chat-people-dialog
				?open=${this.chatPeople && chat !== null}
				.chat=${chat}
				.contacts=${s.contacts}
				@teams-chat-rename=${(e: CustomEvent<{ id: string; name: string }>) => c?.renameChat(e.detail.id, e.detail.name)}
				@teams-chat-add=${(e: CustomEvent<{ id: string; userIds: string[] }>) => c?.addToChat(e.detail.id, e.detail.userIds)}
				@teams-chat-remove=${(e: CustomEvent<{ id: string; userId: string }>) => c?.removeFromChat(e.detail.id, e.detail.userId)}
				@teams-chat-leave=${(e: CustomEvent<{ id: string }>) => {
					if (globalThis.confirm?.(this.t('chatPeople.confirmLeave')) === false) return;
					if (c?.leaveChat(e.detail.id)) {
						this.chatPeople = false;
						this.mobileView = 'list';
					}
				}}
				@teams-dialog-close=${() => (this.chatPeople = false)}
			></teams-chat-people-dialog>
		`;
	}

	// ---- calendar ---------------------------------------------------------------------------
	private calendarMain(s: TeamsState) {
		return html`
			<div class="calendar-back">
				<button type="button" @click=${() => (this.mobileView = 'list')}>
					${icon('back')} ${this.t('app.upcoming')}
				</button>
			</div>
			<office-ui-calendar-view
				.meetings=${s.meetings}
				.zone=${s.viewerZone}
				@office-calendar-open=${(e: CustomEvent<{ meetingId: string; start: number }>) => (this.meetingDetails = { meetingId: e.detail.meetingId, start: e.detail.start })}
				@office-calendar-new=${(e: CustomEvent<{ start: number }>) => (this.meetingForm = { meetingId: '', start: e.detail.start })}
			></office-ui-calendar-view>
		`;
	}

	/** Open the calendar on a meeting's details, for example from its reminder. */
	private openMeeting(meetingId: string, start = 0): void {
		if (!this.closePreview()) return;
		this.settingsOpen = false;
		this.meeting = false;
		this.rail = 'calendar';
		this.mobileView = 'conversation';
		this.meetingDetails = { meetingId, start };
	}

	private async joinMeeting(meetingId: string): Promise<void> {
		if (!this.closePreview()) return;
		this.meetingDetails = null;
		this.settingsOpen = false;
		if (await this.teams.client?.joinMeeting(meetingId)) {
			this.meeting = true;
			this.rail = 'calendar';
			this.mobileView = 'conversation';
		}
	}

	private saveMeeting(form: MeetingFormInput, id?: string, occurrence?: number): boolean {
		const c = this.client;
		if (!c) return false;
		const saved =
			id && occurrence
				? c.editOccurrence(id, occurrence, {
						start: form.start,
						end: form.end,
						title: form.title,
						...(form.description !== undefined ? { description: form.description } : {}),
					})
				: id
					? c.updateMeeting(id, form)
					: c.createMeeting(form);
		if (!saved) return false;
		this.meetingForm = null;
		return true;
	}

	private calendarDialogs(s: TeamsState) {
		const c = this.teams.client;
		const editing = this.meetingForm?.meetingId
			? (s.meetings.find((m) => m.id === this.meetingForm?.meetingId) ?? null)
			: null;
		const details = this.meetingDetails
			? (s.meetings.find((m) => m.id === this.meetingDetails?.meetingId) ?? null)
			: null;
		return html`
			<teams-meeting-dialog
				?open=${this.meetingForm !== null}
				.meeting=${editing}
				.start=${this.meetingForm?.start ?? 0}
				.occurrence=${this.meetingForm?.occurrence ?? 0}
				.contacts=${s.contacts}
				.channels=${s.channels.map((ch) => ({ id: ch.id, name: ch.name }))}
				.save=${(form: MeetingFormInput, id?: string, occurrence?: number) => this.saveMeeting(form, id, occurrence)}
				@teams-dialog-close=${() => (this.meetingForm = null)}
			></teams-meeting-dialog>
			<teams-meeting-details
				?open=${details !== null}
				.meeting=${details}
				.occurrenceStart=${this.meetingDetails?.start ?? 0}
				.zone=${s.viewerZone}
				.channelName=${s.channels.find((ch) => ch.id === details?.channelId)?.name ?? ''}
				@teams-meeting-join=${(e: CustomEvent<{ meetingId: string }>) => void this.joinMeeting(e.detail.meetingId)}
				@teams-meeting-chat=${(e: CustomEvent<{ meetingId: string }>) => {
					const target = s.meetings.find((m) => m.id === e.detail.meetingId);
					const conversation = target?.chatId ?? target?.channelId;
					this.meetingDetails = null;
					if (conversation) this.selectChannel(conversation);
				}}
				@teams-meeting-edit=${(e: CustomEvent<MeetingChange>) => {
					const { meetingId, scope, recurrenceId } = e.detail;
					const shown = this.meetingDetails?.start ?? 0;
					this.meetingDetails = null;
					this.meetingForm =
						scope === 'occurrence'
							? { meetingId, start: shown, occurrence: recurrenceId }
							: { meetingId, start: 0 };
				}}
				@teams-meeting-cancel=${(e: CustomEvent<MeetingChange>) => {
					const { meetingId, scope, recurrenceId } = e.detail;
					if (scope === 'occurrence') {
						if (c?.cancelOccurrence(meetingId, recurrenceId)) this.meetingDetails = null;
						return;
					}
					const ask =
						details?.repeat === 'none'
							? this.t('details.confirmCancel')
							: this.t('details.confirmCancelSeries');
					if (globalThis.confirm?.(ask) === false) return;
					if (c?.cancelMeeting(meetingId)) this.meetingDetails = null;
				}}
				@teams-meeting-ics=${(e: CustomEvent<{ meetingId: string }>) => {
					const file = c?.meetingIcs(e.detail.meetingId);
					if (file && !downloadText(file.name, file.text))
						this.notify(this.t('app.downloadFailed'));
				}}
				@teams-dialog-close=${() => (this.meetingDetails = null)}
			></teams-meeting-details>
		`;
	}

	private async resumeDraft(context: DraftContext): Promise<void> {
		if (!this.closePreview() || !this.teams.client?.openDraft(context)) return;
		this.rail = 'teams';
		this.tab = 'posts';
		this.meeting = false;
		await Promise.resolve();
		await this.updateComplete;
		this.shadowRoot
			?.querySelector('office-ui-chat-composer')
			?.shadowRoot?.querySelector('textarea')
			?.focus();
	}
	private async closeThread(): Promise<void> {
		const client = this.teams.client;
		const root = client?.getState().thread?.root.id;
		client?.closeThread();
		await Promise.resolve();
		await this.updateComplete;
		const tabButton = this.shadowRoot?.querySelector<HTMLElement>('[data-tab-conversation]');
		if (tabButton) {
			tabButton.focus();
			return;
		}
		const list = this.shadowRoot?.querySelector('.conversation-main office-ui-chat-list');
		const message = Array.from(
			list?.shadowRoot?.querySelectorAll<HTMLElement>('[data-message-id]') ?? [],
		).find((element) => element.dataset.messageId === root);
		message?.querySelector<HTMLElement>('.thread-link, [data-action="reply"]')?.focus();
	}

	private pinFile(file: TeamsState['files'][number]): void {
		if (!this.selectChannel(file.channelId ?? this.teams.state?.selectedChannelId ?? '')) return;
		const tab = this.teams.client?.addTab(
			file.name,
			{ type: 'file', attachment: file },
			{ postToChannel: true },
		);
		if (tab) this.selectTab(tab.id);
		else this.notify(this.t('app.pinNeedsUrl'));
	}

	private saveCopyFor(detail: OpenFileDetail, state: TeamsState): SaveFileCopy | undefined {
		const client = this.teams.client;
		const channelId = detail.channelId;
		if (!client || !channelId || !state.canUploadFiles) return undefined;
		return async (file, options) => {
			await client.saveFileCopy(channelId, file, options);
		};
	}

	private fileList(files: TeamsState['files'], state: TeamsState) {
		return html`<teams-files-panel
			.client=${this.teams.client}
			.files=${files}
			.channelId=${state.selectedChannelId}
			.channelName=${state.channel?.name ?? ''}
			.canUpload=${state.canUploadFiles}
			@teams-files-open=${(event: CustomEvent<{ attachment: OpenFileDetail['attachment']; mode?: 'teams' | 'browser' }>) => void this.openFile(event.detail.attachment, event.detail.mode)}
			@teams-files-browser=${(event: CustomEvent<{ attachment: OpenFileDetail['attachment'] }>) => void this.openFile(event.detail.attachment, 'browser')}
			@teams-files-pin=${(event: CustomEvent<{ attachment: TeamsState['files'][number] }>) => this.pinFile(event.detail.attachment)}
		></teams-files-panel>`;
	}

	private allFiles(s: TeamsState) {
		return html`<header class="channel-head">
				<h1>${this.t('rail.files')}</h1>
				<span class="topic">${this.t('app.sharedAcross')}</span>
			</header>
			${this.fileList(s.allFiles, s)}`;
	}

	private meetingView(s: TeamsState) {
		const call = s.call!;
		if (call.phase === 'prejoin')
			return html`
				<office-ui-prejoin
					heading=${this.callHeading(call)}
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
					heading=${this.callHeading(call)}
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
					${
						this.panel === 'chat'
							? html`<aside class="panel chat">${this.conversation(s, true)}</aside>`
							: this.panel === 'people'
								? html`<aside class="panel people-panel">
										<h2>${this.t('app.peopleCount', { count: call.participants.length })}</h2>
										<ul class="plain">
											${call.participants.map(
												(p) =>
													html`<li>
														<office-ui-avatar
															size="sm"
															name=${p.name}
															seed=${p.id}
														></office-ui-avatar>
														<span
															>${p.self ? this.t('common.nameYou', { name: p.name }) : p.name}</span
														>
														${p.state.hand ? html`<span title=${this.t('call.handRaised')}>✋</span>` : nothing}
													</li>`,
											)}
										</ul>
									</aside>`
								: nothing
					}
				</div>
			</section>
		`;
	}
}

export function defineTeamsApp(
	registry: CustomElementRegistry | undefined = globalThis.customElements,
): void {
	if (registry && !registry.get('teams-app')) registry.define('teams-app', TeamsApp);
}
