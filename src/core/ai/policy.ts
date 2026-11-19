// The assistant policy a deployment sets: which providers and models are offered, which features
// are on, what context a request may carry and what every request is told. Hosts load the raw
// configuration (the suite reads `ai-config.json` or `globalThis.ooxmlAiConfig`); this module only
// normalises it and answers what it allows. Without a configuration the assistant is off. The
// policy runs in the browser, so it guides an honest client: rules that must survive a modified
// client belong in a gateway. DOM-free.

/** A provider a connection may name: its id, display name, protocol and default endpoint. */
export interface AiProviderSpec {
	id: string;
	name: string;
	/** `offline` marks the on-device preview, which no connection may name. */
	protocol: string;
	endpoint?: string;
	model?: string;
}

/** A connection the deployment provides. */
export interface AiConnection {
	id: string;
	label: string;
	provider: string;
	endpoint: string;
	model: string;
	models: string[];
	/** `user`: each user enters a key; `none`: no key in the browser (a gateway or local server). */
	keyMode: 'user' | 'none';
	headers: Record<string, string>;
	managed: true;
}

export const AI_FEATURES = [
	'editFiles',
	'createFiles',
	'attachFiles',
	'teamsContext',
	'history',
	'versions',
	'sharedChats',
] as const;
export type AiFeature = (typeof AI_FEATURES)[number];

export interface AiPolicy {
	/** Whether the assistant exists at all. Off unless a configuration is provided. */
	enabled: boolean;
	connections: AiConnection[];
	/** The connection selected for a user who has not picked one (`offline` or a connection id). */
	defaultConnection: string | null;
	/** Whether users may add their own models and keys. */
	allowUserConnections: boolean;
	/** Providers users may add, when they may add connections; every provider when null. */
	userProviders: string[] | null;
	offline: { enabled: boolean };
	features: Record<AiFeature, boolean>;
	/** Where shared chats sync; without a `syncUrl` they use the Teams server or the browser's tabs. */
	collaboration: { syncUrl: string; token: string };
	limits: { maxContextChars: number; historyMessages: number; maxOutputTokens: number };
	/** What a request may read beyond what the user attached. */
	context: {
		/** Hidden sheets, rows and columns go into a request only when this is true. */
		includeHidden: boolean;
	};
	behaviour: {
		instructions: string;
		style: 'concise' | 'balanced' | 'detailed';
		language: string;
		temperature: number | null;
	};
	/** Which behaviour settings a user may change for themselves. */
	userControls: { style: boolean; language: boolean; instructions: boolean; temperature: boolean };
	/** Extra text shown in the model menu, for example who to ask for access. */
	notice: string;
}

function defaults(): AiPolicy {
	return {
		enabled: false,
		connections: [],
		defaultConnection: null,
		allowUserConnections: true,
		userProviders: null,
		offline: { enabled: true },
		features: Object.fromEntries(AI_FEATURES.map((name) => [name, true])) as Record<
			AiFeature,
			boolean
		>,
		collaboration: { syncUrl: '', token: '' },
		limits: { maxContextChars: 150000, historyMessages: 20, maxOutputTokens: 8192 },
		context: { includeHidden: false },
		behaviour: { instructions: '', style: 'balanced', language: '', temperature: null },
		userControls: { style: true, language: true, instructions: true, temperature: false },
		notice: '',
	};
}

type Raw = Record<string, unknown>;
const record = (value: unknown): Raw =>
	value && typeof value === 'object' && !Array.isArray(value) ? (value as Raw) : {};
const bool = (value: unknown, fallback: boolean): boolean =>
	typeof value === 'boolean' ? value : fallback;
const text = (value: unknown, max: number, fallback = ''): string =>
	typeof value === 'string' ? value.slice(0, max) : fallback;
const number = <T extends number | null>(value: unknown, min: number, max: number, fallback: T) =>
	typeof value === 'number' && Number.isFinite(value)
		? Math.min(max, Math.max(min, value))
		: fallback;

function connection(raw: unknown, providers: readonly AiProviderSpec[]): AiConnection | null {
	const r = record(raw);
	const spec = providers.find((p) => p.id === r.provider);
	if (!spec || spec.protocol === 'offline' || !/^[\w-]{1,64}$/u.test(String(r.id ?? '')))
		return null;
	const models = Array.isArray(r.models)
		? r.models.filter((m): m is string => typeof m === 'string' && m !== '').slice(0, 50)
		: [];
	return {
		id: String(r.id),
		label: text(r.label, 80, spec.name),
		provider: spec.id,
		endpoint: text(r.endpoint, 500, spec.endpoint ?? ''),
		model: text(r.model, 200, models[0] ?? spec.model ?? ''),
		models,
		keyMode: r.keyMode === 'none' ? 'none' : 'user',
		headers: Object.fromEntries(
			Object.entries(record(r.headers))
				.filter((entry): entry is [string, string] => {
					const [key, value] = entry;
					return /^[\w-]{1,64}$/u.test(key) && typeof value === 'string';
				})
				.slice(0, 10),
		),
		managed: true,
	};
}

/**
 * A deployment's configuration merged over the defaults, ignoring anything malformed field by
 * field. `providers` is the host's provider catalogue; a connection naming another is dropped.
 * Any configuration turns the assistant on unless it says `"enabled": false`.
 */
export function normalizeAiPolicy(raw: unknown, providers: readonly AiProviderSpec[]): AiPolicy {
	const p = defaults();
	if (!raw || typeof raw !== 'object') return p;
	const r = raw as Raw;
	p.enabled = bool(r.enabled, true);
	p.connections = (Array.isArray(r.connections) ? r.connections : [])
		.map((c) => connection(c, providers))
		.filter((c): c is AiConnection => c !== null)
		.slice(0, 50);
	p.allowUserConnections = bool(r.allowUserConnections, p.allowUserConnections);
	if (Array.isArray(r.userProviders))
		p.userProviders = r.userProviders.filter(
			(id): id is string => typeof id === 'string' && providers.some((s) => s.id === id),
		);
	p.offline.enabled = bool(record(r.offline).enabled, p.offline.enabled);
	const features = record(r.features);
	for (const name of AI_FEATURES) p.features[name] = bool(features[name], p.features[name]);
	const limits = record(r.limits);
	p.limits.maxContextChars = number(limits.maxContextChars, 1000, 1_000_000, 150000);
	p.limits.historyMessages = number(limits.historyMessages, 0, 100, 20);
	p.limits.maxOutputTokens = number(limits.maxOutputTokens, 256, 200_000, 8192);
	p.context.includeHidden = bool(record(r.context).includeHidden, false);
	const behaviour = record(r.behaviour);
	p.behaviour.instructions = text(behaviour.instructions, 8000);
	if (
		behaviour.style === 'concise' ||
		behaviour.style === 'balanced' ||
		behaviour.style === 'detailed'
	)
		p.behaviour.style = behaviour.style;
	p.behaviour.language = text(behaviour.language, 60);
	p.behaviour.temperature = number(behaviour.temperature, 0, 2, null);
	const controls = record(r.userControls);
	for (const key of Object.keys(p.userControls) as (keyof AiPolicy['userControls'])[])
		p.userControls[key] = bool(controls[key], p.userControls[key]);
	p.notice = text(r.notice, 300);
	const collaboration = record(r.collaboration);
	const sync = text(collaboration.syncUrl, 500);
	if (/^wss?:\/\/\S+$/iu.test(sync)) {
		p.collaboration.syncUrl = sync;
		p.collaboration.token = text(collaboration.token, 4096);
	}
	const ids = ['offline', ...p.connections.map((c) => c.id)];
	p.defaultConnection =
		typeof r.defaultConnection === 'string' && ids.includes(r.defaultConnection)
			? r.defaultConnection
			: (p.connections[0]?.id ?? (p.offline.enabled ? 'offline' : null));
	return p;
}

/** Whether `feature` is on: the assistant is enabled and the policy does not turn it off. */
export const aiFeatureEnabled = (policy: AiPolicy, feature: AiFeature): boolean =>
	policy.enabled && policy.features[feature];

/**
 * The endpoint a user may connect to, normalised: HTTPS, or plain HTTP to this computer only,
 * without credentials or a fragment. Throws otherwise.
 */
export function checkAiEndpoint(value: string): string {
	const url = new URL(value);
	if (
		url.username ||
		url.password ||
		url.hash ||
		(url.protocol !== 'https:' &&
			!(url.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname)))
	)
		throw new Error('Use HTTPS, or a local endpoint on localhost.');
	return url.href;
}

/**
 * Why `policy` does not allow a connection the user added (saved earlier, or being saved now),
 * or null when it does: users must be allowed to add connections, its provider must be one the
 * policy lists, and its endpoint must pass {@link checkAiEndpoint}.
 */
export function userConnectionRefusal(
	policy: AiPolicy,
	connection: { provider: unknown; endpoint: unknown },
	providers: readonly AiProviderSpec[],
): string | null {
	if (!policy.enabled) return 'The assistant is off in this workspace.';
	if (!policy.allowUserConnections) return 'This workspace does not allow adding models.';
	const spec = providers.find((p) => p.id === connection.provider);
	if (!spec || spec.protocol === 'offline') return 'That provider is not available.';
	if (policy.userProviders && !policy.userProviders.includes(spec.id))
		return `${spec.name} is not one of the providers this workspace allows.`;
	try {
		checkAiEndpoint(String(connection.endpoint));
	} catch (error) {
		return (error as Error).message;
	}
	return null;
}
