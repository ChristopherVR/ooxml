/**
 * The models a user can pick: the offline preview (if the policy allows it), the connections the
 * deployment provides and, when allowed, connections the user added. Requests go straight from
 * this page to the chosen endpoint. Keys a user enters stay in this tab's session storage unless
 * they ask to remember them on this device; managed gateways need no key in the browser at all.
 */
import { currentProfile } from '../profile-state.js';
import { offlineReply } from './offline.js';
import { currentPolicy } from './policy.js';
import { authHeaders, provider } from './provider-catalog.js';

export { providers, provider } from './provider-catalog.js';

const stateKey = () => `ooxml-suite-ai-v2-${currentProfile().id}`;
const keysKey = () => `ooxml-suite-ai-keys-v2-${currentProfile().id}`;
const listeners = new Set();
const offline = {
	id: 'offline',
	label: 'Offline preview',
	provider: 'offline',
	model: 'Offline preview',
	keyMode: 'none',
	managed: true,
};

function readJson(storage, key) {
	try {
		return JSON.parse(storage.getItem(key)) ?? undefined;
	} catch {
		return undefined;
	}
}
function writeJson(storage, key, value) {
	try {
		storage.setItem(key, JSON.stringify(value));
	} catch {}
}
function state() {
	const saved = readJson(localStorage, stateKey());
	if (saved)
		return { active: saved.active ?? null, user: Array.isArray(saved.user) ? saved.user : [] };
	// One connection saved by the first version of the assistant becomes a user connection.
	const v1 =
		readJson(localStorage, `ooxml-suite-ai-v1-${currentProfile().id}`) ??
		readJson(sessionStorage, `ooxml-suite-ai-v1-${currentProfile().id}`);
	if (v1?.provider && v1.provider !== 'offline' && v1.endpoint) {
		const id = crypto.randomUUID();
		const user = [
			{
				id,
				label: provider(v1.provider).name,
				provider: v1.provider,
				endpoint: v1.endpoint,
				model: v1.model,
			},
		];
		if (v1.key) storeKey(id, v1.key, Boolean(v1.remember));
		writeJson(localStorage, stateKey(), { active: id, user });
		return { active: id, user };
	}
	return { active: null, user: [] };
}
const saveState = (next) => {
	writeJson(localStorage, stateKey(), next);
	for (const listener of listeners) listener();
};
function keyFor(id) {
	return readJson(sessionStorage, keysKey())?.[id] ?? readJson(localStorage, keysKey())?.[id] ?? '';
}
function storeKey(id, key, remember) {
	for (const storage of [sessionStorage, localStorage]) {
		const keys = readJson(storage, keysKey()) ?? {};
		delete keys[id];
		if (key && (storage === localStorage) === remember) keys[id] = key;
		writeJson(storage, keysKey(), keys);
	}
}

/** Every model the user may pick, in menu order. */
export function connections() {
	const policy = currentPolicy();
	return [
		...(policy.offline.enabled ? [offline] : []),
		...policy.connections,
		...(policy.allowUserConnections
			? state().user.map((c) => ({
					...c,
					models: [],
					// A user's own model takes the key they entered with it, if any; it never prompts.
					keyMode: 'own',
					managed: false,
				}))
			: []),
	];
}
/** The selected model with its key, or null when the policy leaves nothing to pick. */
export function currentConnection() {
	const all = connections();
	const { active } = state();
	const chosen =
		all.find((c) => c.id === active) ??
		all.find((c) => c.id === currentPolicy().defaultConnection) ??
		all[0];
	return chosen ? { ...chosen, key: keyFor(chosen.id) } : null;
}
export const needsKey = (c) =>
	Boolean(c && c.keyMode === 'user' && provider(c.provider).auth !== 'none' && !c.key);
export function connectionLabel(c = currentConnection()) {
	if (!c) return 'No model available';
	return c.provider === 'offline' ? c.label : `${c.label} · ${chosenModel(c)}`;
}
export function selectConnection(id, model) {
	const next = state();
	next.active = id;
	if (model) {
		const own = next.user.find((c) => c.id === id);
		if (own) own.model = model;
		else writeJson(localStorage, `${stateKey()}-model-${id}`, model);
	}
	saveState(next);
}
/** The model chosen within a managed connection that offers several. */
export function chosenModel(c) {
	const picked = readJson(localStorage, `${stateKey()}-model-${c.id}`);
	return c.models?.includes(picked) ? picked : c.model;
}
export function setKey(id, key, remember) {
	storeKey(id, key, remember);
	for (const listener of listeners) listener();
}
export const hasRememberedKey = (id) => Boolean(readJson(localStorage, keysKey())?.[id]);
export function saveUserConnection(
	{ id, label, provider: providerId, endpoint, model },
	key,
	remember,
) {
	const next = state();
	const connection = {
		id: id ?? crypto.randomUUID(),
		label,
		provider: providerId,
		endpoint,
		model,
	};
	next.user = [...next.user.filter((c) => c.id !== connection.id), connection];
	next.active = connection.id;
	if (key !== undefined) storeKey(connection.id, key, remember);
	saveState(next);
	return connection;
}
export function removeUserConnection(id) {
	const next = state();
	next.user = next.user.filter((c) => c.id !== id);
	if (next.active === id) next.active = null;
	storeKey(id, '', false);
	saveState(next);
}
export function onConnection(listener) {
	listeners.add(listener);
	return () => listeners.delete(listener);
}
export function checkEndpoint(value) {
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

async function* lines(response) {
	const reader = response.body.getReader();
	const decoder = new TextDecoder();
	let buffer = '';
	for (;;) {
		const { value, done } = await reader.read();
		buffer += decoder.decode(value ?? new Uint8Array(), { stream: !done });
		const parts = buffer.split(/\r?\n/u);
		buffer = done ? '' : parts.pop();
		for (const line of parts) yield line;
		if (done) return;
	}
}
async function failure(response) {
	let detail = '';
	try {
		const body = await response.json();
		detail = body?.error?.message ?? body?.message ?? '';
	} catch {}
	return new Error(
		`The provider returned HTTP ${response.status}${detail ? `: ${String(detail).slice(0, 300)}` : '.'}`,
	);
}

/**
 * Sends one chat turn and streams the reply through `onText` (the text so far). Providers that
 * answer with a single JSON body instead of a stream are read the same way.
 * @returns {Promise<string>} the complete reply
 */
export async function streamChat({
	connection,
	system,
	messages,
	signal,
	onText,
	context,
	settings,
}) {
	const spec = provider(connection.provider);
	if (spec.protocol === 'offline')
		return offlineReply({
			messages,
			context,
			signal,
			onText,
			animate: settings.animate,
			canCreate: settings.canCreate,
			canEdit: settings.canEdit,
		});
	const anthropic = spec.protocol === 'anthropic';
	const model = chosenModel(connection);
	const temperature = settings.temperature ?? undefined;
	const response = await fetch(connection.endpoint, {
		method: 'POST',
		credentials: 'omit',
		redirect: 'error',
		signal: AbortSignal.any([signal, AbortSignal.timeout(180000)]),
		headers: {
			'Content-Type': 'application/json',
			...(connection.headers ?? {}),
			...authHeaders(spec, connection.key),
		},
		body: JSON.stringify(
			anthropic
				? { model, max_tokens: settings.maxTokens, system, messages, stream: true, temperature }
				: {
						model,
						messages: [{ role: 'system', content: system }, ...messages],
						stream: true,
						max_tokens: settings.maxTokens,
						temperature,
					},
		),
	});
	if (!response.ok) throw await failure(response);
	let text = '';
	if (!response.headers.get('content-type')?.includes('text/event-stream')) {
		const body = await response.json();
		text = anthropic
			? (body.content ?? []).map((part) => part.text ?? '').join('')
			: body.choices?.[0]?.message?.content;
		if (typeof text !== 'string') throw new Error('The provider did not return a text response.');
		onText(text);
		return text;
	}
	for await (const line of lines(response)) {
		if (!line.startsWith('data:')) continue;
		const data = line.slice(5).trim();
		if (!data || data === '[DONE]') continue;
		let event;
		try {
			event = JSON.parse(data);
		} catch {
			continue;
		}
		if (event.type === 'error' || event.error)
			throw new Error(event.error?.message ?? 'The provider reported an error.');
		const delta = anthropic
			? event.type === 'content_block_delta'
				? event.delta?.text
				: undefined
			: event.choices?.[0]?.delta?.content;
		if (typeof delta === 'string' && delta) {
			text += delta;
			onText(text);
		}
	}
	return text;
}
