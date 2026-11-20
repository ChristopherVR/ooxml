// Connects a Lit element to a core `TeamsClient`: creates it, re-renders the host whenever the
// state changes and tears it down. This is the whole "framework glue" of the element; the bindings
// for React, Vue, Solid, Svelte and Angular do the same few lines with their own reactivity.
import type { ReactiveController, ReactiveControllerHost } from 'lit';
import {
	type ActivityItem,
	DEFAULT_CHANNEL,
	type TeamsClient,
	type TeamsClientOptions,
	type TeamsState,
	createTeamsClient,
} from 'ooxml-core/teams';
import type { TeamsI18n } from '../i18n/translator';

/**
 * `state` with the workspace's starting channel in the display language while its name and topic
 * are still the ones it was created with (they are stored in English for every peer).
 */
export function localizeDefaultChannel(state: TeamsState, i18n: TeamsI18n): TeamsState {
	if (i18n.locale === 'en') return state;
	const shown = <T extends { id: string; name: string; topic: string }>(channel: T): T => {
		if (channel.id !== DEFAULT_CHANNEL.id) return channel;
		const name =
			channel.name === DEFAULT_CHANNEL.name ? i18n.t('channel.defaultName') : channel.name;
		const topic =
			channel.topic === DEFAULT_CHANNEL.topic ? i18n.t('channel.defaultTopic') : channel.topic;
		return name === channel.name && topic === channel.topic ? channel : { ...channel, name, topic };
	};
	const channels = state.channels.map(shown);
	const channel = state.channel && shown(state.channel);
	if (channel === state.channel && channels.every((entry, i) => entry === state.channels[i]))
		return state;
	return { ...state, channels, channel };
}

export class TeamsController implements ReactiveController {
	client: TeamsClient | null = null;
	private raw: TeamsState | null = null;
	private shown: { raw: TeamsState; i18n: TeamsI18n; state: TeamsState } | undefined;
	private off: (() => void)[] = [];
	private readonly flush = (): void => this.client?.flushStorage();

	constructor(
		private readonly host: ReactiveControllerHost,
		private readonly onNotice: (text: string) => void,
		private readonly onNotification: (item: ActivityItem) => void = () => undefined,
		private readonly i18n?: () => TeamsI18n,
	) {
		host.addController(this);
	}

	/** The client's state as the host shows it (see `localizeDefaultChannel`). */
	get state(): TeamsState | null {
		const raw = this.raw;
		const i18n = this.i18n?.();
		if (!raw || !i18n) return raw;
		if (this.shown?.raw !== raw || this.shown.i18n !== i18n)
			this.shown = { raw, i18n, state: localizeDefaultChannel(raw, i18n) };
		return this.shown.state;
	}

	start(options: TeamsClientOptions): void {
		this.stop();
		const client = createTeamsClient(options);
		this.client = client;
		this.raw = client.getState();
		this.off = [
			client.subscribe(() => {
				this.raw = client.getState();
				this.host.requestUpdate();
			}),
			client.on('notice', this.onNotice),
			client.on('notification', this.onNotification),
		];
		this.host.requestUpdate();
	}

	stop(): void {
		for (const fn of this.off.splice(0)) fn();
		this.client?.destroy();
		this.client = null;
		this.raw = null;
	}

	hostDisconnected(): void {
		globalThis.removeEventListener('pagehide', this.flush);
		this.stop();
	}

	hostConnected(): void {
		globalThis.addEventListener('pagehide', this.flush);
	}
}
