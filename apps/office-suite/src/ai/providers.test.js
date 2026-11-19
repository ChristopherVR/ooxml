import { test } from 'node:test';
import assert from 'node:assert/strict';

// The policy and the saved connections live in browser storage; give the modules a stand-in.
const saved = new Map();
const storage = {
	getItem: (k) => saved.get(k) ?? null,
	setItem: (k, v) => saved.set(k, String(v)),
	removeItem: (k) => saved.delete(k),
};
globalThis.localStorage = storage;
globalThis.sessionStorage = { getItem: () => null, setItem() {}, removeItem() {} };
const { currentProfile } = await import('../profile-state.js');
const { loadPolicy } = await import('./policy.js');
const providers = await import('./providers.js');

test('a saved connection is not offered once userProviders leaves its provider out', async (t) => {
	t.after(() => {
		delete globalThis.ooxmlAiConfig;
		delete globalThis.localStorage;
		delete globalThis.sessionStorage;
	});
	saved.set(
		`ooxml-suite-ai-v2-${currentProfile().id}`,
		JSON.stringify({
			active: 'mine',
			user: [
				{
					id: 'mine',
					label: 'Old',
					provider: 'openrouter',
					endpoint: 'https://openrouter.ai/api/v1/chat/completions',
					model: 'm',
				},
				{
					id: 'tampered',
					label: 'Plain HTTP',
					provider: 'anthropic',
					endpoint: 'http://attacker.test/v1/messages',
					model: 'm',
				},
				{
					id: 'ok',
					label: 'Mine',
					provider: 'anthropic',
					endpoint: 'https://api.anthropic.com/v1/messages',
					model: 'm',
				},
			],
		}),
	);
	globalThis.ooxmlAiConfig = { allowUserConnections: true, userProviders: ['anthropic'] };
	await loadPolicy();
	assert.deepEqual(
		providers.connections().map((c) => c.id),
		['offline', 'ok'],
	);
	assert.notEqual(providers.currentConnection()?.id, 'mine');
	assert.throws(
		() =>
			providers.saveUserConnection({
				label: 'New',
				provider: 'openrouter',
				endpoint: 'https://openrouter.ai/api/v1/chat/completions',
				model: 'm',
			}),
		/OpenRouter is not one of the providers this workspace allows/,
	);
});
