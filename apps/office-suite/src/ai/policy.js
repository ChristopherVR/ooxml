/**
 * The assistant's policy: what a deployment allows, decided outside the UI. An organisation puts
 * an `ai-config.json` beside the app (or a host such as the desktop shell sets
 * `globalThis.ooxmlAiConfig` before the app loads) to choose the providers and models, hold the
 * keys behind its own gateway, turn features off, cap context and set the instructions every
 * request carries. The UI only offers what the policy leaves to the user. Without a file the
 * defaults below apply, so the app works out of the box. See docs/suite-ai-config.md.
 */
import { providers } from './provider-catalog.js';

const defaults = {
	/** Connections the deployment provides; users pick among them. See normalizeConnection. */
	connections: [],
	/** The connection selected for a user who has not picked one ('offline' or a connection id). */
	defaultConnection: null,
	/** Whether users may add their own models and keys from the UI. */
	allowUserConnections: true,
	/** Providers users may add when they may add connections; all of them when omitted. */
	userProviders: null,
	offline: { enabled: true },
	features: {
		editFiles: true,
		createFiles: true,
		attachFiles: true,
		teamsContext: true,
		history: true,
		versions: true,
	},
	limits: { maxContextChars: 150000, historyMessages: 20, maxOutputTokens: 8192 },
	/** Request settings; `instructions` is added to the suite's own system prompt. */
	behaviour: { instructions: '', style: 'balanced', language: '', temperature: null },
	/** Which behaviour settings a user may change for themselves. */
	userControls: { style: true, language: true, instructions: true, temperature: false },
	/** Extra text shown in the model menu, for example who to ask for access. */
	notice: '',
};

let policy = structuredClone(defaults);
let source = 'default';
let loading;

const bool = (value, fallback) => (typeof value === 'boolean' ? value : fallback);
const text = (value, max, fallback = '') =>
	typeof value === 'string' ? value.slice(0, max) : fallback;
const number = (value, min, max, fallback) =>
	typeof value === 'number' && Number.isFinite(value)
		? Math.min(max, Math.max(min, value))
		: fallback;

function normalizeConnection(raw) {
	const spec = providers.find((p) => p.id === raw?.provider);
	if (!spec || spec.protocol === 'offline' || !/^[\w-]{1,64}$/u.test(String(raw.id ?? '')))
		return null;
	const models = Array.isArray(raw.models)
		? raw.models.filter((m) => typeof m === 'string' && m).slice(0, 50)
		: [];
	return {
		id: raw.id,
		label: text(raw.label, 80, spec.name),
		provider: spec.id,
		endpoint: text(raw.endpoint, 500, spec.endpoint ?? ''),
		model: text(raw.model, 200, models[0] ?? spec.model ?? ''),
		models,
		/** 'user': each user enters a key; 'none': no key (a gateway adds it, or a local server). */
		keyMode: raw.keyMode === 'none' ? 'none' : 'user',
		headers:
			raw.headers && typeof raw.headers === 'object'
				? Object.fromEntries(
						Object.entries(raw.headers)
							.filter(([k, v]) => /^[\w-]{1,64}$/u.test(k) && typeof v === 'string')
							.slice(0, 10),
					)
				: {},
		managed: true,
	};
}

/** Merges a deployment's configuration over the defaults, ignoring anything malformed. */
export function normalizePolicy(raw = {}) {
	const p = structuredClone(defaults);
	if (!raw || typeof raw !== 'object') return p;
	p.connections = (Array.isArray(raw.connections) ? raw.connections : [])
		.map(normalizeConnection)
		.filter(Boolean)
		.slice(0, 50);
	p.allowUserConnections = bool(raw.allowUserConnections, p.allowUserConnections);
	if (Array.isArray(raw.userProviders))
		p.userProviders = raw.userProviders.filter((id) => providers.some((s) => s.id === id));
	p.offline.enabled = bool(raw.offline?.enabled, p.offline.enabled);
	for (const key of Object.keys(p.features))
		p.features[key] = bool(raw.features?.[key], p.features[key]);
	p.limits.maxContextChars = number(
		raw.limits?.maxContextChars,
		1000,
		1_000_000,
		p.limits.maxContextChars,
	);
	p.limits.historyMessages = number(raw.limits?.historyMessages, 0, 100, p.limits.historyMessages);
	p.limits.maxOutputTokens = number(
		raw.limits?.maxOutputTokens,
		256,
		200_000,
		p.limits.maxOutputTokens,
	);
	p.behaviour.instructions = text(raw.behaviour?.instructions, 8000);
	p.behaviour.style = ['concise', 'balanced', 'detailed'].includes(raw.behaviour?.style)
		? raw.behaviour.style
		: p.behaviour.style;
	p.behaviour.language = text(raw.behaviour?.language, 60);
	p.behaviour.temperature = number(raw.behaviour?.temperature, 0, 2, null);
	for (const key of Object.keys(p.userControls))
		p.userControls[key] = bool(raw.userControls?.[key], p.userControls[key]);
	p.notice = text(raw.notice, 300);
	const ids = ['offline', ...p.connections.map((c) => c.id)];
	p.defaultConnection = ids.includes(raw.defaultConnection)
		? raw.defaultConnection
		: (p.connections[0]?.id ?? (p.offline.enabled ? 'offline' : null));
	return p;
}

/** Loads the deployment's policy once: a host-provided object, else `ai-config.json`. */
export function loadPolicy() {
	loading ??= (async () => {
		let raw = globalThis.ooxmlAiConfig;
		if (raw) source = 'host';
		else
			try {
				const root = document.querySelector('meta[name=ooxml-root]')?.content ?? './';
				const response = await fetch(new URL(`${root}ai-config.json`, location.href), {
					cache: 'no-store',
					credentials: 'same-origin',
				});
				if (response.ok && response.headers.get('content-type')?.includes('json')) {
					raw = await response.json();
					// The shipped file says "managed": false; an organisation's own file manages the assistant.
					source = raw?.managed === false ? 'default' : 'file';
				}
			} catch {}
		policy = normalizePolicy(raw);
		return policy;
	})();
	return loading;
}

export const currentPolicy = () => policy;
/** Whether an organisation configured the assistant (so the UI can say who manages it). */
export const isManaged = () => source !== 'default';
export const feature = (name) => policy.features[name] !== false;
