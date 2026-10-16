/**
 * The providers the assistant can talk to. Most speak the OpenAI chat-completions protocol;
 * Anthropic has its own. `auth` says how a key is sent. Requests go straight from the browser,
 * so a provider must allow cross-origin calls; one that does not is reached through an
 * organisation gateway, configured as a `gateway` (or `custom`) connection that adds the key on
 * the server. Model names are left to the deployment or the user where they change often.
 */
export const providers = [
	{
		id: 'offline',
		name: 'Offline preview',
		protocol: 'offline',
		note: 'No model. Simple rules on this device show how chat works and build sample files.',
	},
	{
		id: 'anthropic',
		name: 'Anthropic',
		protocol: 'anthropic',
		endpoint: 'https://api.anthropic.com/v1/messages',
		model: 'claude-sonnet-5-5',
		auth: 'x-api-key',
	},
	{
		id: 'openai',
		name: 'OpenAI',
		protocol: 'openai',
		endpoint: 'https://api.openai.com/v1/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'github',
		name: 'GitHub Models',
		protocol: 'openai',
		endpoint: 'https://models.github.ai/inference/chat/completions',
		auth: 'bearer',
		keyLabel: 'GitHub token (models:read)',
	},
	{
		id: 'azure',
		name: 'Azure OpenAI',
		protocol: 'openai',
		endpoint:
			'https://YOUR-RESOURCE.openai.azure.com/openai/deployments/YOUR-DEPLOYMENT/chat/completions?api-version=2024-10-21',
		auth: 'api-key',
		note: 'Use your resource and deployment in the endpoint; the model is the deployment name.',
	},
	{
		id: 'google',
		name: 'Google Gemini',
		protocol: 'openai',
		endpoint: 'https://generativelanguage.googleapis.com/v1beta/openai/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'mistral',
		name: 'Mistral',
		protocol: 'openai',
		endpoint: 'https://api.mistral.ai/v1/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'groq',
		name: 'Groq',
		protocol: 'openai',
		endpoint: 'https://api.groq.com/openai/v1/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'deepseek',
		name: 'DeepSeek',
		protocol: 'openai',
		endpoint: 'https://api.deepseek.com/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'xai',
		name: 'xAI',
		protocol: 'openai',
		endpoint: 'https://api.x.ai/v1/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'together',
		name: 'Together AI',
		protocol: 'openai',
		endpoint: 'https://api.together.xyz/v1/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'fireworks',
		name: 'Fireworks AI',
		protocol: 'openai',
		endpoint: 'https://api.fireworks.ai/inference/v1/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'perplexity',
		name: 'Perplexity',
		protocol: 'openai',
		endpoint: 'https://api.perplexity.ai/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'openrouter',
		name: 'OpenRouter',
		protocol: 'openai',
		endpoint: 'https://openrouter.ai/api/v1/chat/completions',
		auth: 'bearer',
	},
	{
		id: 'ollama',
		name: 'Ollama on this computer',
		protocol: 'openai',
		endpoint: 'http://localhost:11434/v1/chat/completions',
		model: 'llama3.2',
		auth: 'none',
	},
	{
		id: 'lmstudio',
		name: 'LM Studio on this computer',
		protocol: 'openai',
		endpoint: 'http://localhost:1234/v1/chat/completions',
		auth: 'none',
	},
	{
		id: 'gateway',
		name: 'Organisation gateway',
		protocol: 'openai',
		endpoint: '',
		auth: 'none',
		note: 'An OpenAI-compatible endpoint run by your organisation that adds credentials itself.',
	},
	{
		id: 'custom',
		name: 'Custom (OpenAI-compatible)',
		protocol: 'openai',
		endpoint: '',
		auth: 'bearer',
	},
];

export const provider = (id) => providers.find((p) => p.id === id) ?? providers[0];

/** The request headers that carry a key for this provider. */
export function authHeaders(spec, key) {
	if (!key || spec.auth === 'none') return {};
	if (spec.auth === 'x-api-key')
		return {
			'x-api-key': key,
			'anthropic-version': '2023-06-01',
			'anthropic-dangerous-direct-browser-access': 'true',
		};
	if (spec.auth === 'api-key') return { 'api-key': key };
	return { Authorization: `Bearer ${key}` };
}
