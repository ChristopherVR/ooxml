/**
 * Model providers for the assistant. Every request goes straight from this page to the provider
 * the user configured; there is no suite server in between. Keys stay in this tab's session
 * storage unless the user asks to remember them on this device.
 */
import { currentProfile } from '../profile-state.js';
import { offlineReply } from './offline.js';

export const providers = [
	{
		id: 'offline',
		name: 'Offline preview',
		protocol: 'offline',
		note: 'No model. Simple rules on this device show how chat works and build sample files.',
	},
	{
		id: 'openai',
		name: 'OpenAI',
		protocol: 'openai',
		endpoint: 'https://api.openai.com/v1/chat/completions',
		model: '',
		key: true,
	},
	{
		id: 'anthropic',
		name: 'Anthropic',
		protocol: 'anthropic',
		endpoint: 'https://api.anthropic.com/v1/messages',
		model: 'claude-sonnet-5-5',
		key: true,
	},
	{
		id: 'openrouter',
		name: 'OpenRouter',
		protocol: 'openai',
		endpoint: 'https://openrouter.ai/api/v1/chat/completions',
		model: '',
		key: true,
	},
	{
		id: 'ollama',
		name: 'Ollama on this computer',
		protocol: 'openai',
		endpoint: 'http://localhost:11434/v1/chat/completions',
		model: 'llama3.2',
		key: false,
	},
	{
		id: 'custom',
		name: 'Custom (OpenAI-compatible)',
		protocol: 'openai',
		endpoint: '',
		model: '',
		key: true,
	},
];
export const provider = (id) => providers.find((p) => p.id === id) ?? providers[0];

const storageKey = () => `ooxml-suite-ai-v1-${currentProfile().id}`;
const listeners = new Set();
function read(storage) {
	try {
		return JSON.parse(storage.getItem(storageKey())) ?? undefined;
	} catch {
		return undefined;
	}
}
/** The active connection: the offline preview until the user connects a model. */
export function currentConnection() {
	const saved = read(localStorage) ?? read(sessionStorage);
	return saved?.provider && provider(saved.provider).id === saved.provider
		? saved
		: { provider: 'offline', model: 'Offline preview' };
}
export function connectionLabel(connection = currentConnection()) {
	return connection.provider === 'offline'
		? 'Offline preview'
		: `${provider(connection.provider).name} · ${connection.model}`;
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
export function saveConnection(next, remember) {
	const value = JSON.stringify(next);
	for (const storage of [localStorage, sessionStorage])
		try {
			storage.removeItem(storageKey());
		} catch {}
	try {
		(remember ? localStorage : sessionStorage).setItem(storageKey(), value);
	} catch {}
	for (const listener of listeners) listener(next);
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
export async function streamChat({ connection, system, messages, signal, onText, context }) {
	const spec = provider(connection.provider);
	if (spec.protocol === 'offline') return offlineReply({ messages, context, signal, onText });
	const anthropic = spec.protocol === 'anthropic';
	const response = await fetch(connection.endpoint, {
		method: 'POST',
		credentials: 'omit',
		redirect: 'error',
		signal: AbortSignal.any([signal, AbortSignal.timeout(180000)]),
		headers: {
			'Content-Type': 'application/json',
			...(anthropic
				? {
						'x-api-key': connection.key ?? '',
						'anthropic-version': '2023-06-01',
						'anthropic-dangerous-direct-browser-access': 'true',
					}
				: connection.key
					? { Authorization: `Bearer ${connection.key}` }
					: {}),
		},
		body: JSON.stringify(
			anthropic
				? { model: connection.model, max_tokens: 8192, system, messages, stream: true }
				: {
						model: connection.model,
						messages: [{ role: 'system', content: system }, ...messages],
						stream: true,
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
