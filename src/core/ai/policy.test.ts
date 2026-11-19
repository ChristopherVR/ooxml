import { describe, expect, test } from 'vitest';
import {
	aiFeatureEnabled,
	checkAiEndpoint,
	normalizeAiPolicy,
	userConnectionRefusal,
	type AiProviderSpec,
} from './policy';

const providers: AiProviderSpec[] = [
	{ id: 'offline', name: 'Offline preview', protocol: 'offline' },
	{ id: 'anthropic', name: 'Anthropic', protocol: 'anthropic', endpoint: 'https://a.test/v1' },
	{ id: 'openrouter', name: 'OpenRouter', protocol: 'openai', endpoint: 'https://o.test/v1' },
	{ id: 'ollama', name: 'Ollama', protocol: 'openai', endpoint: 'http://localhost:11434/v1' },
];

describe('normalizeAiPolicy', () => {
	test('without a configuration the assistant is off', () => {
		for (const raw of [undefined, null, 'on', 42]) {
			const policy = normalizeAiPolicy(raw, providers);
			expect(policy.enabled).toBe(false);
			expect(aiFeatureEnabled(policy, 'editFiles')).toBe(false);
		}
	});

	test('any configuration turns it on unless it says otherwise', () => {
		expect(normalizeAiPolicy({}, providers).enabled).toBe(true);
		expect(normalizeAiPolicy({ enabled: false }, providers).enabled).toBe(false);
	});

	test('malformed fields fall back one by one', () => {
		const policy = normalizeAiPolicy(
			{
				connections: [
					{ id: 'team', provider: 'anthropic', models: ['m1', 2, 'm2'] },
					{ id: 'bad id!', provider: 'anthropic' },
					{ id: 'unknown', provider: 'nowhere' },
					{ id: 'local', provider: 'offline' },
				],
				userProviders: ['anthropic', 'nowhere', 7],
				features: { editFiles: 'no', createFiles: false },
				limits: { maxContextChars: 5, historyMessages: 'many' },
				behaviour: { style: 'shouty', temperature: 9 },
				collaboration: { syncUrl: 'http://not-a-socket', token: 'secret' },
				defaultConnection: 'missing',
			},
			providers,
		);
		expect(policy.connections.map((c) => [c.id, c.model, c.models])).toEqual([
			['team', 'm1', ['m1', 'm2']],
		]);
		expect(policy.userProviders).toEqual(['anthropic']);
		expect(policy.features.editFiles).toBe(true);
		expect(policy.features.createFiles).toBe(false);
		expect(policy.limits).toEqual({
			maxContextChars: 1000,
			historyMessages: 20,
			maxOutputTokens: 8192,
		});
		expect(policy.behaviour).toMatchObject({ style: 'balanced', temperature: 2 });
		expect(policy.collaboration).toEqual({ syncUrl: '', token: '' });
		expect(policy.defaultConnection).toBe('team');
	});

	test('hidden workbook content stays out of requests unless the policy opts in', () => {
		expect(normalizeAiPolicy({}, providers).context.includeHidden).toBe(false);
		expect(normalizeAiPolicy({ context: { includeHidden: 'yes' } }, providers).context).toEqual({
			includeHidden: false,
		});
		expect(normalizeAiPolicy({ context: { includeHidden: true } }, providers).context).toEqual({
			includeHidden: true,
		});
	});
});

describe('userConnectionRefusal', () => {
	const openrouter = { provider: 'openrouter', endpoint: 'https://o.test/v1' };
	test.each([
		['no configuration', undefined, openrouter, /assistant is off/],
		['users may not add models', { allowUserConnections: false }, openrouter, /does not allow/],
		['provider outside userProviders', { userProviders: ['anthropic'] }, openrouter, /OpenRouter/],
		['unknown provider', {}, { provider: 'nowhere', endpoint: 'https://x.test' }, /not available/],
		['the offline preview', {}, { provider: 'offline', endpoint: '' }, /not available/],
		[
			'plain HTTP to another host',
			{},
			{ provider: 'openrouter', endpoint: 'http://o.test' },
			/HTTPS/,
		],
		[
			'credentials in the URL',
			{},
			{ provider: 'openrouter', endpoint: 'https://u:p@o.test' },
			/HTTPS/,
		],
		['a malformed endpoint', {}, { provider: 'openrouter', endpoint: 'not a url' }, /URL/i],
	])('refuses %s', (_, raw, connection, reason) => {
		expect(userConnectionRefusal(normalizeAiPolicy(raw, providers), connection, providers)).toMatch(
			reason,
		);
	});

	test.each([
		['every provider by default', {}, openrouter],
		[
			'a listed provider',
			{ userProviders: ['anthropic'] },
			{ provider: 'anthropic', endpoint: 'https://a.test' },
		],
		['a local server over HTTP', {}, { provider: 'ollama', endpoint: 'http://localhost:11434/v1' }],
	])('allows %s', (_, raw, connection) => {
		expect(
			userConnectionRefusal(normalizeAiPolicy(raw, providers), connection, providers),
		).toBeNull();
	});
});

test('checkAiEndpoint normalises an allowed endpoint', () => {
	expect(checkAiEndpoint('https://api.example.test/v1')).toBe('https://api.example.test/v1');
	expect(() => checkAiEndpoint('https://api.example.test/v1#x')).toThrow(/HTTPS/);
});
